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
    public partial class frmSeekFor : Form
    {
        private const string TABANA = "AnaFornitori";

        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        public string _strCod = "";
        public string _strDes = "";
        public bool _bolNew = true;
        public string _strSqlWhe = "";
        public bool _bolAnagra = true;
        private string _strConSql = "";

        public DataTable _tabTmp = new DataTable();

        public frmSeekFor()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
        }

        private void frmSeekFor_Load(object sender, EventArgs e)
        {
            ApplyModernUi();
            _strConSql = _clsFun.ConSql("");

            SetDgv1();
            if (!_bolAnagra && !_bolNew)
                btnNew.Enabled = false;
            txtDes.Select();

            clsUiIcons.RestoreFormBounds(this);
            clsUiIcons.RestoreGridColumnWidths(dgv1, "frmSeekFor_dgv1");
        }

        private void frmSeekFor_FormClosing(object sender, FormClosingEventArgs e)
        {
            clsUiIcons.SaveFormBounds(this);
            clsUiIcons.SaveGridColumnWidths(dgv1, "frmSeekFor_dgv1");
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

                if (dgv1 != null)
                {
                    clsUiIcons.StyleDataGridView(dgv1);
                }
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("frmSeekFor.ApplyModernUi", ex.Message);
            }
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }

        private void frmSeekFor_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
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

        private void txtDes_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
                TabSeek("DES");
        }

        private void SetDgv1()
        {
            dgv1.AutoGenerateColumns = false;
            dgv1.AllowUserToAddRows = false;
            dgv1.ReadOnly = false;
            dgv1.AllowUserToDeleteRows = false;

            DataGridViewTextBoxColumn cTbc;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "for_cod";
            cTbc.Name = "Codice";
            cTbc.Width = 140;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "for_des";
            cTbc.Name = "Descrizione";
            cTbc.Width = 400;
            cTbc.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);
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
                    w = "for_des LIKE '%" + txtDes.Text.Trim() + "%'";
            }
            else if (strTip == "COD")
            {
                if (txtCod.Text != "")
                {
                    txtCod.Text = txtCod.Text.PadLeft(5, Convert.ToChar("0"));
                    w = "for_cod='" + txtCod.Text + "'";
                }
            }
            if (w != "")
            {
                if (!chkAnn.Checked)
                    w += " AND for_ann=0 ";

                _strSqlWhe = s;
                if (strTip == "ALL")
                {
                    s = "SELECT * FROM " + p + " ";
                    s += "WHERE ";     
                    if (!chkAnn.Checked)
                        s += "for_ann=0 ";
                    s += "ORDER BY for_des";
                }
                else
                {
                    //w += " AND art_cli='" + _clsDef.COD06X + "' ";

                    s = "SELECT * FROM " + p + " WHERE ";
                    s += w;
                    s += " ORDER BY for_des";
                }
                DataTable t = _clsFun.FillTabSql(TABANA, s, false, _strConSql);
                t.DefaultView.AllowDelete = false;
                t.DefaultView.AllowEdit = false;
                t.DefaultView.AllowNew = false;
                dgv1.DataSource = t;
                dgv1.Focus();
            }
            else if (strTip != "")
            {
                MessageBox.Show("Fornitore non trovato");
            }
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            frmAnaFornitore f = new frmAnaFornitore();
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
            Scelto();
        }

        private void dgv1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
                Scelto();
        }

        private void Scelto()
        {
            CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, this.dgv1.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {
                DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                DataRow x = r.Row;

                _strCod = Convert.ToString(x["for_cod"]);
                _strDes = Convert.ToString(x["for_des"]);

                _tabTmp = new clsGenTabTmp().TabTmpFor("AnaFor");
                DataRow k = _tabTmp.NewRow();
                k["tmp_for"] = _strCod;
                k["tmp_fod"] = _strDes;
                _tabTmp.Rows.Add(k);

                if (_bolAnagra)
                {
                    frmAnaFornitore f = new frmAnaFornitore();
                    f._tabTmp = _tabTmp.Copy();
                    f.ShowDialog();
                    if (f._bolCodNew && f._strCod != _clsDef.CODNEW)
                        TabSeek("DES");
                    x["for_des"] = f._strDes;
                }
                else
                {
                    if ((Boolean)x["for_ann"])
                        MessageBox.Show("Fornitore annullato");
                    else
                        Esci();
                }
            }
            else
                Esci();
        }

    }
}
