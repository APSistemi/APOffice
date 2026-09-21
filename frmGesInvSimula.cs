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
    public partial class frmGesInvSimula : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        public decimal _decVal = 0;
        public decimal _decArt = 0;

        private string _strConSql = "";

        public frmGesInvSimula()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0); ;
        }

        private void frmGesInvSimula_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }
        private void frmGesInvSimula_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
        }
        private void Esci()
        {
            this.Close();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            if (_clsFun.Numerico(txtImp.Text))
                _decVal = Convert.ToDecimal(txtImp.Text);
            if (_clsFun.Numerico(txtArt.Text))
                _decArt = Convert.ToDecimal(txtArt.Text);

            if (_decVal == 0 || _decArt == 0)
                MessageBox.Show("Importo obbiettivo o numero articoli non definiti correttamente!", "Simulazione inventario", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
                Esci();
        }

    }
}
