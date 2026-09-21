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
    public partial class frmUtyConf : Form
    {
        public frmUtyConf()
        {
            InitializeComponent();
        }

        private void frmConfUty_Load(object sender, EventArgs e)
        {

        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            new frmUtyConf1().ShowDialog();
        }

    }
}
