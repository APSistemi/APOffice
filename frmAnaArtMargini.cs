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
    public partial class frmAnaArtMargini : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        public decimal _decCos = 0;
        public decimal _decPrv = 0;
        public decimal _decSfr = 0;
        public decimal _decIva = 0;
        public string _strLis = "";

        public DataTable _tabLiv = new DataTable();

        public frmAnaArtMargini()
        {
            InitializeComponent();
        }

        private void frmAnaArtMargini_Load(object sender, EventArgs e)
        {
            ApplyModernUi();
            clsUiIcons.RestoreFormBounds(this);

            lblPrc.Text = _decCos.ToString("#,##0.000");
            txtPrc.Text = lblPrc.Text;

            lblPrv.Text = _decPrv.ToString("#,##0.00");
            txtPrv.Text = lblPrv.Text;

            lblMav.Text = _clsFun.Dec2Txt(_clsFun.Margine(_decPrv, _decCos, _decIva, _decSfr, "P"), 2);
            txtMav.Text = lblMav.Text;

            lblRic.Text = _clsFun.Dec2Txt(_clsFun.Ricarico(_decPrv, _decCos, _decIva, "P"), 2);
            txtRic.Text = lblRic.Text;

            FillTab();

            this.ActiveControl = txtPrv;
            txtPrv.SelectAll();
        }

        private void frmAnaArtMargini_FormClosing(object sender, FormClosingEventArgs e)
        {
            clsUiIcons.SaveFormBounds(this);
        }

        private void ApplyModernUi()
        {
            try
            {
                this.BackColor = Color.FromArgb(243, 244, 246);

                // MenuStrip Moderno
                if (menuStrip1 != null)
                {
                    menuStrip1.Renderer = clsUiIcons.GetModernMenuRenderer();
                    if (esciToolStripMenuItem != null)
                        esciToolStripMenuItem.Image = clsUiIcons.GetIcon("exit", 16);
                }

                // Pulsante Conferma (Emerald Green Glossy)
                clsUiIcons.StyleStatButton(btnOk, "CONFERMA (F5 / Invio)", "check", 16,
                    Color.FromArgb(236, 253, 245), Color.FromArgb(167, 243, 208), Color.FromArgb(52, 211, 153),
                    Color.FromArgb(6, 78, 59), Color.FromArgb(5, 150, 105));

                // Pulsante Annulla (Soft Red Glossy)
                clsUiIcons.StyleStatButton(btnAnnulla, "Annulla (Esc)", "exit", 14,
                    Color.FromArgb(254, 242, 242), Color.FromArgb(254, 202, 202), Color.FromArgb(248, 113, 113),
                    Color.FromArgb(127, 29, 29), Color.FromArgb(220, 38, 38));
            }
            catch { }
        }

        private void FillTab()
        {
            if (_tabLiv != null)
            {
                cmbLisVen.DataSource = _tabLiv;
                cmbLisVen.DisplayMember = "tab_des";
                cmbLisVen.ValueMember = "tab_cod";
                if (!string.IsNullOrEmpty(_strLis))
                    cmbLisVen.SelectedValue = _strLis;
            }
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }

        private void frmAnaArtMargini_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                Esci();
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.F5)
            {
                btnOk_Click(null, null);
                e.Handled = true;
            }
        }

        private void Esci()
        {
            this.Close();
        }

        private void txtVal_Enter(object sender, EventArgs e)
        {
            if (sender is TextBox txt)
            {
                txt.SelectAll();
            }
        }

        private void txtVal_Leave(object sender, EventArgs e)
        {
            if (sender is TextBox txt)
            {
                CalcolaValori(txt);
            }
        }

        private void txtVal_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                if (sender is TextBox txt)
                {
                    CalcolaValori(txt);
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                }
            }
        }

        private void CalcolaValori(TextBox txt)
        {
            if (txt == null) return;

            decimal dPrc = 0;
            decimal dPrv = 0;
            decimal dMar = 0;
            decimal dRic = 0;

            if (_clsFun.Numerico(txtPrc.Text, "0123456789,."))
                dPrc = Convert.ToDecimal(txtPrc.Text.Replace(".", ","));
            if (_clsFun.Numerico(txtPrv.Text, "0123456789,."))
                dPrv = Convert.ToDecimal(txtPrv.Text.Replace(".", ","));
            if (_clsFun.Numerico(txtMav.Text, "0123456789,."))
                dMar = Convert.ToDecimal(txtMav.Text.Replace(".", ","));
            if (_clsFun.Numerico(txtRic.Text, "0123456789,."))
                dRic = Convert.ToDecimal(txtRic.Text.Replace(".", ","));

            if (txt.Name == "txtPrc")
            {
                dPrv = _clsFun.Marg2Prv(dMar, dPrc, _decIva, _decSfr);
                dMar = _clsFun.Margine(dPrv, dPrc, _decIva, _decSfr, "P");
                dRic = _clsFun.Ricarico(dPrv, dPrc, _decIva, "P");
            }
            else if (txt.Name == "txtPrv")
            {
                dMar = _clsFun.Margine(dPrv, dPrc, _decIva, _decSfr, "P");
                dRic = _clsFun.Ricarico(dPrv, dPrc, _decIva, "P");
            }
            else if (txt.Name == "txtMav")
            {
                dPrv = _clsFun.Marg2Prv(dMar, dPrc, _decIva, _decSfr);
                dRic = _clsFun.Ricarico(dPrv, dPrc, _decIva, "P");
            }
            else if (txt.Name == "txtRic")
            {
                decimal k = (dRic >= 5.0m) ? (1.0m + (dRic / 100.0m)) : ((dRic >= 1.0m) ? dRic : (1.0m + dRic));
                decimal pNetto = dPrc * k;
                dPrv = _clsFun.PiuIva(pNetto, _decIva);
                dMar = _clsFun.Margine(dPrv, dPrc, _decIva, _decSfr, "P");
            }

            txtPrc.Text = dPrc.ToString("#0.000");
            txtPrv.Text = dPrv.ToString("#0.00");
            txtMav.Text = dMar.ToString("#0.00");
            txtRic.Text = dRic.ToString("#0.00");
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            if (this.ActiveControl is TextBox txt)
            {
                CalcolaValori(txt);
            }

            _decPrv = 0;
            if (_clsFun.Numerico(txtPrv.Text, "0123456789,."))
                _decPrv = Convert.ToDecimal(txtPrv.Text.Replace(".", ","));

            if (cmbLisVen.SelectedValue != null && !string.IsNullOrEmpty(cmbLisVen.SelectedValue.ToString()))
                _strLis = cmbLisVen.SelectedValue.ToString();
            if (string.IsNullOrEmpty(_strLis))
                _strLis = _clsDef.LISPOS;

            this.DialogResult = DialogResult.OK;
            Esci();
        }

        private void btnAnnulla_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            Esci();
        }
    }
}
