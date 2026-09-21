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
    public partial class frmUtyPeriodo : Form
    {
        public string _strRes = "";
        public DateTime _dayIni = new DateTime();
        public DateTime _dayFin = new DateTime();

        public frmUtyPeriodo()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0); ;
        }

        private void frmUtyPeriodo_Load(object sender, EventArgs e)
        {
            dtpIni.Value = _dayIni;
            dtpFin.Value = _dayFin;
        }

        private void abbandonaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }
        private void confermaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //SendKeys.Send("{ENTER}");

            Application.DoEvents();

            string s = "";
            s = dtpIni.Text;
            s = dtpFin.Text;

            Console.WriteLine("xxxxxxx");

            _dayIni = dtpIni.Value;
            _dayFin = dtpFin.Value;
            _strRes = "S";

            Esci();
        }
        private void frmUtyPeriodo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
        }
        private void Esci()
        {
            this.Close();
        }

    }
}
