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
    public partial class frmMnuUtility : Form
    {

        public frmMnuUtility()
        {
            InitializeComponent();
        }

        private void frmMnUtility_Load(object sender, EventArgs e)
        {

        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }

        private void frmMnUtility_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
        }

        private void Esci()
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            new frmUtyApShopAttArts().ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            new frmUtyPluEanAttivazione().ShowDialog();
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            new frmUtyArfErrato().ShowDialog();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            new frmUtyAggPrzAnaArticoli().ShowDialog();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            new frmUtyArtScontoVenduto().ShowDialog();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            new frmUtyIva().ShowDialog();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            new frmUtyCtrlArtStat().ShowDialog();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            new frmUtyEstrazioni().ShowDialog();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            new frmUtyCtrlArtPos().ShowDialog();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            new frmUtyCtrlPrv().ShowDialog();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            new frmUtyCtrlJournal().ShowDialog();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            new frmUtyArtMerceologie().ShowDialog();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            new frmUtyArtFillMerceologie().ShowDialog();
        }

        private void button14_Click(object sender, EventArgs e)
        {
            new frmUtyCtrlLisFor().ShowDialog();
        }
    }
}
