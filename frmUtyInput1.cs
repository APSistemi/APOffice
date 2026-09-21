using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace APOffice
{
    public partial class frmUtyInput1 : Form
    {
        public string _strDes = "";
        public string _strTxt = "";

        public frmUtyInput1()
        {
            InitializeComponent();
        }

        private void frmUtyInput1_Load(object sender, EventArgs e)
        {
            if (_strDes.Length > 20)
                lblMsg.Font = new Font("Arial", 10);
            lblMsg.Text = _strDes;
            txtTxt.MaxLength = 500;
            txtTxt.Text = _strTxt;
            txtTxt.SelectionStart = txtTxt.Text.Length;
        }

        private void btnBrowse_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "Database MS Access (*.mdb)|*.mdb|Tutti i file (*.*)|*.*";
            ofd.Title = "Seleziona file Database MDB";
            if (!string.IsNullOrEmpty(txtTxt.Text) && File.Exists(txtTxt.Text))
            {
                try { ofd.InitialDirectory = Path.GetDirectoryName(txtTxt.Text); } catch { }
            }
            else
            {
                ofd.InitialDirectory = @"C:\ApProject\ApShop\DataBase\";
            }

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                txtTxt.Text = ofd.FileName;
            }
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            Esci();
        }

        private void frmUtyInput1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
                Esci();
            else if (e.KeyCode == Keys.Escape)
            {
                this.Close();
            }
        }

        private void Esci()
        {
            _strTxt = txtTxt.Text;
            this.Close();
        }
    }
}
