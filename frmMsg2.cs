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
    public partial class frmMsg2 : Form
    {
        public frmMsg2()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);;
        }

        private void frmMsg2_Load(object sender, EventArgs e)
        {

        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
