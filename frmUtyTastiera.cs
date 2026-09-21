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
    public partial class frmUtyTastiera : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        private string _strConMdb = "";
        public string _strStr = "";
        public Boolean _bolPwd = false;

        private const string TABFON = "TabFonts";

        public frmUtyTastiera()
        {
            InitializeComponent(); 
            _strConMdb = _clsFun.ConMdb("");
        }

        private void frmUtyTastiera_Load(object sender, EventArgs e)
        {
            CtrlButtons();
            FillImg();
            txtSeek.Text = _strStr;
            txtSeek.Select();

            if (_bolPwd)
                txtSeek.PasswordChar = '*';
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }

        private void Esci()
        {
            _strStr = txtSeek.Text;
            this.Close();
        }

        private void CtrlButtons()
        {
            Image imgBtn1 = new Bitmap(@"C:\ApProject\Temp\Img\ApShop\btn03.png");

            string s = "SELECT * FROM " + TABFON;
            //DataTable tFon = _clsFun.FillTabMdb(TABFON, s, false, _clsFun.ConMdb(_clsDef.MDBFON));
            //DataRow[] j = tFon.Select("tab_cod='009'");
            //Font fonBtn = _clsFun.FontTab(tFon, "009");

            Font fonBtn = new Font("Times New Roman", 12.0f);

            foreach (Control c in pnlTasti.Controls)
            {
                if (c.GetType() == typeof(Button) && c.Name.Substring(0, 3) == "btn")
                {
                    Button b = (Button)c;
                    b.Visible = true;
                    b.BackColor = Color.Transparent;
                    b.ForeColor = Color.White;
                    b.BackgroundImageLayout = ImageLayout.Stretch;
                    b.FlatStyle = FlatStyle.Flat;
                    b.FlatAppearance.BorderSize = 0;
                    b.BackgroundImage = imgBtn1;
                    b.Font = fonBtn;
                }
            }
        }

        private void FillImg()
        {
            Image imgBtn1 = new Bitmap(@"C:\ApProject\Temp\Img\ApShop\btn03.png");
            Image imgBtn2 = new Bitmap(@"C:\ApProject\Temp\Img\ApShop\btn03.png");
            Image imgCalc = new Bitmap(@"C:\ApProject\Temp\Img\ApShop\btnCal.png");

            string s = "SELECT * FROM " + TABFON;
            //DataTable tFon = _clsFun.FillTabMdb(TABFON, s, false, _clsFun.ConMdb(_clsDef.MDBFON));
            Font font = new Font("Times New Roman", 12.0f);

            foreach (ToolStripMenuItem t in menuStrip1.Items)
            {
                if (t.Text != "")
                {
                    t.BackgroundImage = imgBtn1;
                    t.BackColor = Color.Transparent;
                    t.ForeColor = Color.White;
                    t.BackgroundImageLayout = ImageLayout.Stretch;
                    //toolItem.FlatStyle = FlatStyle.Flat;
                    //toolItem.FlatAppearance.BorderSize = 0;
                    //btn.Font = new System.Drawing.Font("Microsoft Sans Serif", 28.0f, System.Drawing.FontStyle.Bold);
                    //t.Font = _clsFun.FontTab(tFon, "005");
                    t.Font = font;
                }
            }
        }

        private void btn_Click(object sender, EventArgs e)
        {
            Button btn = (Button)sender;

            if (btn.Name.Length > 3)
            {
                string sSeek = txtSeek.Text;

                String sCmd = btn.Name.Substring(3);

                if (sCmd == "Can")
                {
                    if (sSeek.Length > 0)
                        sSeek = sSeek.Substring(0, sSeek.Length - 1);
                }
                else if (sCmd == "Vrg")
                    sSeek = sSeek + ",";
                else if (sCmd == "Pun")
                    sSeek = sSeek + ".";
                else if (sCmd == "Tra")
                    sSeek = sSeek + "-";
                else if (sCmd == "Inv")
                    Esci();
                else
                    sSeek = sSeek + sCmd;

                txtSeek.Text = sSeek;
                txtSeek.Select();
            }
        }

        private void txtSeek_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
                Esci();
        }


    }
}
