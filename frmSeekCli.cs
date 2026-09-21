using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System.IO;
using System.Diagnostics;

namespace APOffice
{
    public partial class frmSeekCli : Form
    {
        private const string TABANA = "AnaClienti";
        private const string TABTABSTD = "TabStatoDivulgazioni";

        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        public string _strCod = "";
        public string _strDes = "";
        public bool _bolNew = true;
        public string _strSqlWhe = "";
        public bool _bolAnagra = true;
        public bool _bolSelect = false;
        private string _strConSql = "";

        public DataTable _tabTmp = new DataTable();

        public frmSeekCli()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
        }

        private void frmSeekCli_Load(object sender, EventArgs e)
        {
            ApplyModernUi();
            _strConSql = _clsFun.ConSql("");

            SetDgv1();
            if (!_bolAnagra)
                btnNew.Enabled = false;
            txtDes.Select();

            clsUiIcons.RestoreFormBounds(this);
            clsUiIcons.RestoreGridColumnWidths(dgv1, "frmSeekCli_dgv1");
        }

        private void frmSeekCli_FormClosing(object sender, FormClosingEventArgs e)
        {
            clsUiIcons.SaveFormBounds(this);
            clsUiIcons.SaveGridColumnWidths(dgv1, "frmSeekCli_dgv1");
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

                if (btnAll != null)
                {
                    clsUiIcons.StyleStatButton(btnAll, "Tutti", "list", 18,
                        Color.FromArgb(239, 246, 255), Color.FromArgb(191, 219, 254),
                        Color.FromArgb(96, 165, 250), Color.FromArgb(30, 58, 138), Color.FromArgb(29, 78, 216));
                }

                if (btnNew != null)
                {
                    clsUiIcons.StyleStatButton(btnNew, "Nuovo", "plus", 18,
                        Color.FromArgb(236, 253, 245), Color.FromArgb(167, 243, 208),
                        Color.FromArgb(52, 211, 153), Color.FromArgb(6, 78, 59), Color.FromArgb(5, 150, 105));
                }

                if (btnPrn != null)
                {
                    clsUiIcons.StyleStatButton(btnPrn, "Stampa", "print", 16,
                        Color.FromArgb(254, 252, 232), Color.FromArgb(254, 240, 138),
                        Color.FromArgb(250, 204, 21), Color.FromArgb(113, 63, 18), Color.FromArgb(202, 138, 4));
                }

                if (btnPos != null)
                {
                    clsUiIcons.StyleStatButton(btnPos, "Invio alle casse", "send", 16,
                        Color.FromArgb(240, 253, 250), Color.FromArgb(153, 246, 228),
                        Color.FromArgb(45, 212, 191), Color.FromArgb(19, 78, 74), Color.FromArgb(13, 148, 136));
                }

                if (dgv1 != null)
                {
                    clsUiIcons.StyleDataGridView(dgv1);
                }
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("frmSeekCli.ApplyModernUi", ex.Message);
            }
        }

        private void frmSeekCli_KeyDown(object sender, KeyEventArgs e)
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

        private void btnAll_Click(object sender, EventArgs e)
        {
            TabSeek("ALL");
        }

        private void txtCod_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
                TabSeek("COD");
        }

        private void txtPiv_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
                TabSeek("PIV");
        }

        private void txtDes_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
                TabSeek("DES");
        }

        private void SetDgv1()
        {
            DataTable tSta = _clsFun.FillTabSql(TABTABSTD, "SELECT * FROM TabStatoDivulgazioni WHERE tab_var=1", false, _strConSql);

            dgv1.AutoGenerateColumns = false;
            dgv1.AllowUserToAddRows = false;
            dgv1.ReadOnly = false;
            dgv1.AllowUserToDeleteRows = false;

            DataGridViewTextBoxColumn cTbc;
            DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "cli_cod";
            cTbc.Name = "Codice";
            cTbc.Width = 120;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "cli_des";
            cTbc.Name = "Descrizione";
            cTbc.Width = 350;
            cTbc.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cCmb = new DataGridViewComboBoxColumn();
            cCmb.DataPropertyName = "cli_div";
            cCmb.Name = "Stato";
            cCmb.Width = 120;
            cCmb.DataSource = tSta;
            cCmb.ValueMember = "tab_cod";
            cCmb.DisplayMember = "tab_des";
            cCmb.ReadOnly = true;
            cCmb.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing;
            cCmb.DefaultCellStyle.Font = new Font("Segoe UI", 9F, GraphicsUnit.Point);
            dgv1.Columns.Add((DataGridViewColumn)cCmb);
        }

        private void TabSeek(string strTip)
        {
            string p = TABANA;
            string s = "";
            string w = "";

            if (_strSqlWhe != "" && strTip == "")
            {
                w = _strSqlWhe;
            }
            else if (strTip == "ALL")
            {
                w = strTip;
            }
            else if (strTip == "DES")
            {
                if (txtDes.Text != "")
                    w = "cli_des LIKE '%" + txtDes.Text.Trim() + "%'";
            }
            else if (strTip == "PIV")
            {
                if (txtPiv.Text != "")
                    w = "cli_piv = '" + txtPiv.Text.Trim() + "'";
            }
            else if (strTip == "COD")
            {
                if (txtCod.Text != "")
                {
                    //txtCod.Text = txtCod.Text.PadLeft(5, Convert.ToChar("0"));
                    //w = "cli_cod='" + txtCod.Text + "'";

                    txtCod.Text = txtCod.Text.PadLeft(5, Convert.ToChar("0"));
                    w = "cli_cod='" + txtCod.Text + "'";

                }
            }
            if (w != "")
            {
                if (!chkAnn.Checked)
                    w += " AND cli_ann=0 ";

                _strSqlWhe = s;
                if (strTip == "ALL")
                {
                    s = "SELECT * FROM " + p + " ";
                    s += "WHERE ";
                    if (!chkAnn.Checked)
                        s += "cli_ann=0 ";
                    s += "ORDER BY cli_des";
                }
                else
                {
                    //w += " AND art_cli='" + _clsDef.COD06X + "' ";

                    s = "SELECT * FROM " + p + " WHERE ";
                    s += w;
                    s += " ORDER BY cli_des";
                }
                DataTable t = _clsFun.FillTabSql(TABANA, s, false, _strConSql);

                //t.Columns.Add(new DataColumn()
                //{
                //    DataType = Type.GetType("System.String"),
                //    ColumnName = "CliInv",
                //    Caption = "Invio",
                //    MaxLength = 1,
                //    ReadOnly = false,
                //    DefaultValue = (String)"5"
                //});

                t.DefaultView.AllowDelete = false;
                t.DefaultView.AllowEdit = false;
                t.DefaultView.AllowNew = false;
                dgv1.DataSource = t;
                dgv1.Focus();
            }
            else if (strTip != "")
            {
                MessageBox.Show("Cliente non trovato");
            }
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            frmAnaCliente f = new frmAnaCliente();
            f._strCod = _clsDef.CODNEW;
            f.ShowDialog();

            if (f._strCod != "" && f._strCod != _clsDef.CODNEW)
            {
                txtCod.Text = f._strCod;
                TabSeek("COD");
            }
        }

        private void dgv1_DoubleClick(object sender, EventArgs e)
        {
            Scelto("");
        }

        private void dgv1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 1)
            {
                Scelto("ANAGRA");
            }
            else
                Scelto("");
        }

        private void dgv1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
                Scelto("");
        }

        private void dgv1_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 1)
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    Scelto("ANAGRA");
                }
            }
        }

        private void dgv1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 2)
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    string s = (string)x["cli_div"];
                    if (s == "5" || s == "")
                        x["cli_div"] = "0";
                    else if (s == "0")
                        x["cli_div"] = "5";
                }
            }
        }

        private void chkAll_Click(object sender, EventArgs e)
        {
            string sSta = "5";
            if (chkAll.Checked)
                sSta = "0";
            DataTable t = (DataTable)dgv1.DataSource;
            foreach (DataRow y in t.Rows)
                y["cli_div"] = sSta;
        }

        private void Scelto(string strTip)
        {
            CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, this.dgv1.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {
                DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                DataRow x = r.Row;

                _strCod = Convert.ToString(x["cli_cod"]);
                _strDes = Convert.ToString(x["cli_des"]);

                _tabTmp = new clsGenTabTmp().TabTmpCli("AnaCli");
                DataRow k = _tabTmp.NewRow();
                k["tmp_cli"] = _strCod;
                k["tmp_cld"] = _strDes;

                k["tmp_tpa"] = "";
                if (!DBNull.Value.Equals(x["cli_tpa"]))
                    k["tmp_tpa"] = Convert.ToString(x["cli_tpa"]); 

                _tabTmp.Rows.Add(k);

                if (strTip == "ANAGRA" || (_bolAnagra && !_bolSelect))
                {
                    frmAnaCliente f = new frmAnaCliente();
                    f._tabTmp = _tabTmp.Copy();
                    f.ShowDialog();
                    if (f._bolCodNew && f._strCod != _clsDef.CODNEW)
                        TabSeek("DES");
                    x["cli_des"] = f._strDes;
                    //f.Dispose();
                    //SendKeys.SendWait("=");
                }
                else
                {
                    if ((Boolean)x["cli_ann"])
                        MessageBox.Show("Cliente annullato");
                    else
                        Esci();
                }
            }
            else
                Esci();
        }

        private void btnPrn_Click(object sender, EventArgs e)
        {
            if (dgv1.DataSource == null)
                MessageBox.Show("Non ci sono data da stampare!", "CONTROLLO CLIENTI", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            else
                PdfClienti();
        }

        private void PdfClienti()
        {
            string s = "";
            DataRow[] j;
            int iRows = 37;
            int iRow = 0;

            DataTable t = (DataTable)dgv1.DataSource;

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Clienti";

            // Create an empty page
            PdfPage page = pd.AddPage();

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 10, XFontStyle.Bold);
            XFont font2 = new XFont("Courier new", 9, XFontStyle.Regular);
            XFont font3 = new XFont("Courier new", 8, XFontStyle.Regular);
            XFont font4 = new XFont("Courier new", 8, XFontStyle.Bold);

            double Y = 5;
            double X = 20;

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            string sFam = ""; // (string)t.Rows[0]["InvL1c"] + (string)t.Rows[0]["InvL2c"] + (string)t.Rows[0]["InvL3c"];

            decimal dRig = 0;
            decimal dPez = 0;
            decimal dVal = 0;

            for (int i = 0; i <= t.Rows.Count - 1; i++)
            {
                if (iRow > iRows || iRow == 0)
                {
                    if (iRow > 0)
                    {
                        page = pd.AddPage();
                        gfx = XGraphics.FromPdfPage(page);
                    }

                    X = 20;
                    iRow = 0;
                    s = DateTime.Today.ToString("dd/MM/yyyy") + " STAMPA ELENCO CLIENTI";
                    gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    X += 20;

                    s = "Codice";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 60, X, XStringFormats.Default);

                    s = "Descrizione";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 100, X, XStringFormats.Default);

                    X += 15;
                }

                iRow++;

                if (t.Rows.Count - 1 >= i)
                {

                    s = (string)t.Rows[i]["cli_cod"];
                    gfx.DrawString(s, font4, XBrushes.Black, Y + 60, X, XStringFormats.Default);

                    s = (string)t.Rows[i]["cli_des"];
                    gfx.DrawString(s, font4, XBrushes.Black, Y + 100, X, XStringFormats.Default);

                    s = (string)t.Rows[i]["cli_piv"];
                    gfx.DrawString(s, font4, XBrushes.Black, Y + 180, X, XStringFormats.Default);

                    //break;

                    dRig++;

                }

                X += 20;
                //break;

            }

            X += 25;

            //s = "Totale righe " + lblInvRig.Text;
            //gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

            // Save the document...
            try
            {
                string sFil = "C:\\ApProject\\PDF\\Cli_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                pd.Save(sFil);
                // ...and start a viewer.
                Process.Start(sFil);
            }
            catch (Exception ex)
            {
                MessageBox.Show("File di stampa già aperto!");
            }

        }

        private void btnPos_Click(object sender, EventArgs e)
        {
            if (_clsFun.ParGet(clsDefine.enuParametri.Par023Clienti2Pos, _strConSql) == "S")
            {
                if (dgv1.DataSource == null || ((DataTable)dgv1.DataSource).Rows.Count == 0)
                {
                    MessageBox.Show("Righe non presenti!");
                }
                else
                {
                    if (MessageBox.Show("Confermi l'invio in cassa dei clienti selezionati?", "DIVULGAZIONE CLIENTI", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        Cli2Pos();
                }
            }
            else
                MessageBox.Show("Configurazione non definita!", "DIVULGAZIONE CLIENTI", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void Cli2Pos()
        {
            string sMsg = "";
            string s = "";
            DataRow[] j;

            DataTable t = (DataTable)dgv1.DataSource;
            DataTable t2 = t.Clone();

            foreach(DataRow y in t.Rows)
            {
                if ((string)y["cli_div"] == _clsDef.DIVDAD)
                {
                    if (((string)y["cli_piv"]).Trim() == "" && ((string)y["cli_cfi"]).Trim() == "")
                        sMsg += ((string)y["cli_des"]).Trim() + " con partita IVA e Codice fiscale mancanti " + _clsDef.CRLF;
                    else
                    {
                        s = "SELECT * FROM AnaClienti WHERE cli_cod='" + y["cli_cod"] + "'";
                        DataTable tt = _clsFun.FillTabSql("AnaClienti", s, true, _strConSql);
                        if (tt.Rows.Count > 0)
                        {
                            tt.Rows[0]["cli_div"] = _clsDef.DIVDAD;
                            t2.ImportRow(tt.Rows[0]);
                        }
                    }
                }
            }

            t = t2.Copy();

            if (t.Rows.Count > 0)
            {
                frmGesVarPos f = new frmGesVarPos();
                f._tabCli = t;
                f.ShowDialog();

                foreach (DataRow y in ((DataTable)dgv1.DataSource).Rows)
                {
                    j = t.Select("cli_cod='" + y["cli_cod"] + "'");
                    if (j.Length > 0)
                        y["cli_div"] = "5";
                }
            }

            if (sMsg != "")
                MessageBox.Show(sMsg, "CONTROLLO CLIENTI", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }


    }
}
