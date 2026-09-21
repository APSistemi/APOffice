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
    public partial class frmMsg1 : Form
    {
        public string _strMsg = "";
        public string _strTip = "";
        public string _strRes = "";
        public Boolean _bolMem = false;
        public string _strMem = "";

        public frmMsg1()
        {
            InitializeComponent();
        }

        private void frmMsg1_Load(object sender, EventArgs e)
        {
            lblMsg1.Text = _strMsg;

            if (_strTip == "OK")
            {
                btnNo.Visible = false;
                btnSi.Visible = false;
            }
            else if (_strTip == "SINO")
                btnOk.Visible = false;

            if (!_bolMem)
                chkMem.Visible = false;
        }

        private void btn_Click(object sender, EventArgs e)
        {
            Button btn = (Button)sender;
            _strRes = btn.Text;
            if (chkMem.Checked)
                _strMem = btn.Text;
            this.Close();
        }
    }
}
