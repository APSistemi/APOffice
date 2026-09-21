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
    public partial class frmAnaFornitore : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        private const string TABANA = "AnaFornitori";
        private string _strConSql = "";

        public string _strCod = "";
        public string _strDes = "";
        public DataTable _tabTmp = new DataTable();

        public Boolean _bolCodNew = false;

        public frmAnaFornitore()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
        }

        private void frmAnaFornitore_Load(object sender, EventArgs e)
        {
            ApplyModernUi();
            _strConSql = _clsFun.ConSql("");
            FillData();

            clsUiIcons.RestoreFormBounds(this);
        }

        private void frmAnaFornitore_FormClosing(object sender, FormClosingEventArgs e)
        {
            clsUiIcons.SaveFormBounds(this);
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
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("frmAnaFornitore.ApplyModernUi", ex.Message);
            }
        }

        private void frmAnaFornitore_KeyDown(object sender, KeyEventArgs e)
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
            _strDes = txtForDes.Text;

            if (!b)
            {
                if (MessageBox.Show("Continui l'uscita?", this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                      == DialogResult.Yes)
                    this.Close();
            }
            else this.Close();
        }

        private void FillData()
        {

            if (_strCod == _clsDef.CODNEW)
            {
                txtForCod.Text = _clsDef.CODNEW;
                _strCod = _clsDef.CODNEW;
                //dtpforDti.Value = DateTime.Today;
                //dtpforDtm.Value = DateTime.Today;
            }
            else
            {
                if (_tabTmp.Rows.Count > 0)
                {
                    _strCod = (string)_tabTmp.Rows[0]["tmp_for"];
                }

                txtForCod.Text = _strCod;
                string s = "SELECT * FROM " + TABANA + " WHERE for_cod = '" + _strCod + "'";
                DataTable t = _clsFun.FillTabSql(TABANA, s, true, _strConSql);

                if (t.Rows.Count > 0)
                {
                    txtForCod.Text = (string)t.Rows[0]["for_cod"];
                    txtForDes.Text = (string)t.Rows[0]["for_des"];
                    txtForInd.Text = (string)t.Rows[0]["for_ind"];
                    txtForLoc.Text = (string)t.Rows[0]["for_loc"];
                    txtForCap.Text = (string)t.Rows[0]["for_cap"];
                    txtForPrv.Text = (string)t.Rows[0]["for_prv"];

                    txtForPiv.Text = (string)t.Rows[0]["for_piv"];
                    txtForCfi.Text = (string)t.Rows[0]["for_cfi"];


                    txtForTel.Text = (string)t.Rows[0]["for_tel"];
                    txtForCel.Text = (string)t.Rows[0]["for_cel"];
                    txtForFax.Text = (string)t.Rows[0]["for_fax"];
                    txtForMai.Text = (string)t.Rows[0]["for_mai"];
                    txtForNot.Text = (string)t.Rows[0]["for_not"];
                }
            }
        }

        private Boolean Salva()
        {
            Boolean b = true;
            string s = "";
            DataRow y;

            if (txtForDes.Text == "")
            {
                MessageBox.Show("Nome mancante!");
                return false;
            }

            if (txtForCod.Text == _clsDef.CODNEW)
            {
                txtForCod.Text = _clsFun.NewNum(_clsDef.COD04Z, clsDefine.enuNumeratori.NumAnaFornitori, 5, _strConSql);
                _strCod = txtForCod.Text;
            }

            s = "SELECT * FROM " + TABANA + " WHERE for_cod = '" + txtForCod.Text + "'";

            DataTable t = _clsFun.FillTabSql(TABANA, s, true, _strConSql);

            y = t.NewRow();
            y["for_cod"] = txtForCod.Text;

            y["for_des"] = txtForDes.Text.ToUpper();
            y["for_ind"] = txtForInd.Text.ToUpper();
            y["for_loc"] = txtForLoc.Text.ToUpper();
            y["for_cap"] = txtForCap.Text;
            y["for_prv"] = txtForPrv.Text.ToUpper();

            y["for_piv"] = txtForPiv.Text;
            y["for_cfi"] = txtForCfi.Text;
            y["for_tel"] = txtForTel.Text;
            y["for_cel"] = txtForCel.Text;
            y["for_fax"] = txtForFax.Text;
            y["for_mai"] = txtForMai.Text.ToLower();
            y["for_not"] = txtForNot.Text;

            if (t.Rows.Count == 0)
                s = _clsFun.SqlInsertRow(TABANA, t, y);
            else
            {
                ArrayList aWhe = new ArrayList();
                aWhe.Add("for_cod");
                ArrayList aExl = new ArrayList();
                s = _clsFun.SqlUpdRow(TABANA, t, t.Rows[0], y, aWhe, aExl);
            }

            if (s != "")
            { 
                _clsFun.SqlWrite(s, _strConSql);
                _clsFun.FileLog("AnaFornitore", txtForCod.Text, s);
            }

            return (b);
        }

        private void txtForPiv_Validated(object sender, EventArgs e)
        {
            if(txtForPiv.Text != "")
            {
                if (!new clsCtrlCodici().ControllaPartitaIva(txtForPiv.Text))
                        MessageBox.Show("Partita IVA non corretta!");
            }
        }

        private void txtForCfi_Validated(object sender, EventArgs e)
        {
            if (txtForCfi.Text != "")
            {
                if (new clsCtrlCodici().VerificaCodiceFiscale(txtForCfi.Text))
                    MessageBox.Show("Codice fiscale non corretto!");
            }
        }

    }
}
