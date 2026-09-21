using System;
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
    public partial class frmSeekOff : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        private const string TABANAOFT = "GesOffTestate";
        private const string TABTABSTA = "TabStato";
        
        private string _strConSql = "";

        public frmSeekOff()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
        }

        private void frmSeekOff_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");
            ApplyModernUi();
            SetDgv1();
            FillDati();
            clsUiIcons.RestoreFormBounds(this);
            clsUiIcons.RestoreGridColumnWidths(dgv1, "frmSeekOff_dgv1");
            LayoutPromoButtons();
        }

        private void frmSeekOff_FormClosing(object sender, FormClosingEventArgs e)
        {
            clsUiIcons.SaveFormBounds(this);
            clsUiIcons.SaveGridColumnWidths(dgv1, "frmSeekOff_dgv1");
        }

        private void frmSeekOff_Resize(object sender, EventArgs e)
        {
            LayoutPromoButtons();
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
                    if (utilitàToolStripMenuItem != null)
                        utilitàToolStripMenuItem.Image = clsUiIcons.GetIcon("utilita", 16);
                    if (cancellaToolStripMenuItem != null)
                        cancellaToolStripMenuItem.Image = clsUiIcons.GetIcon("trash", 16);
                }

                this.BackColor = Color.FromArgb(243, 244, 246);
                if (groupBox1 != null)
                {
                    groupBox1.BackColor = Color.FromArgb(248, 250, 252);
                    groupBox1.ForeColor = Color.FromArgb(30, 41, 59);
                    groupBox1.Font = new Font("Segoe UI", 9.5f, FontStyle.Bold);
                }

                // 1. Tasto "Per articolo" -> Gradiente Blu Premium con icona offers_promo
                clsUiIcons.StyleStatButton(btnNew, "Per articolo", "offers_promo", 22,
                    Color.FromArgb(239, 246, 255), Color.FromArgb(191, 219, 254),
                    Color.FromArgb(96, 165, 250), Color.FromArgb(30, 58, 138), Color.FromArgb(29, 78, 216));

                // 2. Tasto "Su totale scontrino" -> Gradiente Smeraldo con icona receipt
                clsUiIcons.StyleStatButton(btnTran, "Su totale scontrino", "receipt", 22,
                    Color.FromArgb(236, 253, 245), Color.FromArgb(167, 243, 208),
                    Color.FromArgb(52, 211, 153), Color.FromArgb(6, 78, 59), Color.FromArgb(5, 150, 105));

                // 3. Tasto "Su paniere" -> Gradiente Ambra Caldo con icona basket
                clsUiIcons.StyleStatButton(button1, "Su paniere", "basket", 22,
                    Color.FromArgb(255, 251, 235), Color.FromArgb(254, 230, 138),
                    Color.FromArgb(251, 191, 36), Color.FromArgb(120, 53, 15), Color.FromArgb(217, 119, 6));

                if (dgv1 != null)
                {
                    clsUiIcons.StyleDataGridView(dgv1);
                    dgv1.CellFormatting += dgv1_CellFormatting;
                }
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("frmSeekOff.ApplyModernUi", ex.Message);
            }
        }

        private void LayoutPromoButtons()
        {
            try
            {
                if (groupBox1 == null || btnNew == null || btnTran == null || button1 == null) return;
                int margin = 10;
                int gap = 12;
                int availableWidth = groupBox1.ClientSize.Width - (margin * 2) - (gap * 2);
                int btnWidth = Math.Max(120, availableWidth / 3);
                int btnHeight = Math.Max(38, groupBox1.ClientSize.Height - 30);
                int btnY = 22;

                btnNew.SetBounds(margin, btnY, btnWidth, btnHeight);
                btnTran.SetBounds(margin + btnWidth + gap, btnY, btnWidth, btnHeight);
                button1.SetBounds(margin + (btnWidth + gap) * 2, btnY, btnWidth, btnHeight);
            }
            catch { }
        }

        private void dgv1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;

            string colName = dgv1.Columns[e.ColumnIndex].Name;

            // Formattazione Date
            if ((colName == "Data inizio" || colName == "Data fine") && e.Value != null)
            {
                if (e.Value is DateTime dt)
                {
                    e.Value = dt.ToString("dd/MM/yyyy");
                    e.FormattingApplied = true;
                }
                else if (DateTime.TryParse(e.Value.ToString(), out DateTime parsedDt))
                {
                    e.Value = parsedDt.ToString("dd/MM/yyyy");
                    e.FormattingApplied = true;
                }
            }

            // Formattazione Badge Colonna Stato
            if (colName == "Stato" && e.Value != null)
            {
                string val = e.Value.ToString().Trim().ToUpperInvariant();
                if (val.Contains("ATTIV") || val.Contains("CORSO") || val == "00")
                {
                    e.CellStyle.BackColor = Color.FromArgb(220, 252, 231);
                    e.CellStyle.ForeColor = Color.FromArgb(22, 101, 52);
                    e.CellStyle.Font = new Font("Segoe UI", 9.0f, FontStyle.Bold);
                }
                else if (val.Contains("PROGRAMM") || val.Contains("DA ATTIVARE"))
                {
                    e.CellStyle.BackColor = Color.FromArgb(224, 242, 254);
                    e.CellStyle.ForeColor = Color.FromArgb(3, 105, 161);
                    e.CellStyle.Font = new Font("Segoe UI", 9.0f, FontStyle.Bold);
                }
                else if (val.Contains("SCADUT") || val.Contains("NON ATTIV"))
                {
                    e.CellStyle.BackColor = Color.FromArgb(254, 243, 199);
                    e.CellStyle.ForeColor = Color.FromArgb(146, 64, 14);
                    e.CellStyle.Font = new Font("Segoe UI", 9.0f, FontStyle.Bold);
                }
                else if (val.Contains("CANCELL") || val.Contains("ELIMIN"))
                {
                    e.CellStyle.BackColor = Color.FromArgb(254, 226, 226);
                    e.CellStyle.ForeColor = Color.FromArgb(153, 27, 27);
                    e.CellStyle.Font = new Font("Segoe UI", 9.0f, FontStyle.Bold);
                }
            }
        }

        private void frmSeekOff_KeyDown(object sender, KeyEventArgs e)
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

        private void SetDgv1()
        {
            dgv1.AutoGenerateColumns = false;
            dgv1.AllowUserToAddRows = false;
            dgv1.ReadOnly = true;
            dgv1.AllowUserToDeleteRows = false;
            dgv1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dgv1.MultiSelect = false;

            DataGridViewTextBoxColumn cTbc;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "oft_yea";
            cTbc.Name = "Anno";
            cTbc.HeaderText = "Anno";
            cTbc.Width = 55;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "oft_cod";
            cTbc.Name = "Codice";
            cTbc.HeaderText = "Codice";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "oft_des";
            cTbc.Name = "Descrizione";
            cTbc.HeaderText = "Descrizione offerta";
            cTbc.Width = 320;
            cTbc.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            cTbc.FillWeight = 150;
            cTbc.MinimumWidth = 180;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "oft_dti";
            cTbc.Name = "Data inizio";
            cTbc.HeaderText = "Data inizio";
            cTbc.Width = 95;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "oft_dtf";
            cTbc.Name = "Data fine";
            cTbc.HeaderText = "Data fine";
            cTbc.Width = 95;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "OftStd";
            cTbc.Name = "Stato";
            cTbc.HeaderText = "Stato";
            cTbc.Width = 115;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);
        }

        private void FillDati()
        {
            DataRow[] j;

            string s = "SELECT ";
            s += "tab_cod, tab_des, tab_art, tab_off ";
            s += "FROM TabStato";
            DataTable tSta = _clsFun.FillTabSql(TABTABSTA, s, false, _strConSql);

            s = "SELECT ";
            s += "GesOffTestate.* ";
            s += "FROM GesOffTestate ";
            s += "ORDER BY oft_dti DESC";
            DataTable t = _clsFun.FillTabSql(TABANAOFT, s, false, _strConSql);

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "OftStd",
                Caption = "Stato",
                MaxLength = 20,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            foreach(DataRow y in t.Rows)
            {
                s = (string)y["oft_cod"];

                if((string)y["oft_sta"] == _clsDef.STANOA)
                {
                    if ((DateTime)y["oft_dti"] <= DateTime.Today && (DateTime)y["oft_dtf"] >= DateTime.Today)
                        y["oft_sta"] = _clsDef.STADAA;
                }

                j = tSta.Select("tab_cod='" + y["oft_sta"] + "'");
                if (j.Length > 0)
                    y["OftStd"] = j[0]["tab_des"];

            }

            t.DefaultView.AllowDelete = false;
            t.DefaultView.AllowEdit = false;
            t.DefaultView.AllowNew = false;
            dgv1.DataSource = t;
            dgv1.Focus();
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            frmGesOfferta f = new frmGesOfferta();
            f._strOftYea = DateTime.Today.Year.ToString();
            f._strOftCod = _clsDef.CODNEW;
            f.ShowDialog();
            FillDati();
        }
        
        private void dgv1_DoubleClick(object sender, EventArgs e)
        {
            CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, this.dgv1.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {
                DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                DataRow x = r.Row;

                frmGesOfferta f = new frmGesOfferta();
                f._strOftYea = (string)x["oft_yea"];
                f._strOftCod = (string)x["oft_cod"];
                f._strOftSta = (string)x["oft_sta"];
                f.ShowDialog();
                if (f._dasGen.Tables.IndexOf(TABANAOFT) >= 0)
                {
                    DataTable t = f._dasGen.Tables[TABANAOFT];
                    if(t.Rows.Count > 0)
                    {
                        x["oft_des"] = t.Rows[0]["oft_des"];
                        x["oft_dti"] = t.Rows[0]["oft_dti"];
                        x["oft_dtf"] = t.Rows[0]["oft_dtf"];
                        x["oft_sta"] = t.Rows[0]["oft_sta"];
                        x["OftStd"] = t.Rows[0]["OftStd"];
                    }
                }
            }
        }

        private void btnTran_Click(object sender, EventArgs e)
        {
            new frmGesOffTransazione().ShowDialog();
        }

        private void cancellaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CanOfferta();
        }

        private void CanOfferta()
        {
            CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {
                DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                DataRow x = r.Row;

                string s = (string)x["oft_sta"];

                string sYea = (string)x["oft_yea"];
                string sCod = (string)x["oft_cod"];

                if(s != _clsDef.STACAN)
                    MessageBox.Show("Stato offerta non è CANCELLATO!", "CONTROLLO CANCELLAZIONE", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                else
                {
                    if (MessageBox.Show("Cancellazione dell'offerta " + sCod + "/" + sYea + ", confermi?", "CANCELLAZIONE OFFERTA", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    {
                        s = "DELETE FROM GesOffArticoli WHERE ofa_yea='" + sYea + "' AND ofa_cod='" + sCod + "'";
                        _clsFun.SqlWrite(s, _strConSql);

                        s = "DELETE FROM GesOffTestate WHERE oft_yea='" + sYea + "' AND oft_cod='" + sCod + "'";
                        _clsFun.SqlWrite(s, _strConSql);
                    }

                    FillDati();
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {

        }

    }
}
