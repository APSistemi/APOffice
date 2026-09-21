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
    public partial class frmGesVarPosImpianto : Form
    {
        public Boolean _bolOk = false;
        public Boolean _bolAll = false;
        
        public frmGesVarPosImpianto()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
        }

        private void frmGesVarPosImpianto_Load(object sender, EventArgs e)
        {
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }

        private void Esci()
        {
            this.Close();
        }

        private void btnFill_Click(object sender, EventArgs e)
        {
            _bolOk = true;
            if (rdbAll.Checked)
                _bolAll = true;
            Esci();
        }

    }
}
