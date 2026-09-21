using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace APOffice
{
    public partial class frmMnuUtilita : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        private string _strConSql = "";

        public frmMnuUtilita()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
            SetIcons();
        }

        private void frmMnuUtilita_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");
            panel1.Select();
            ApplyModernUi();
        }

        private void SetIcons()
        {
            // Colonna 1
            btnConf.Image = GetIcon("⚙️", Color.DimGray);
            button43.Image = GetIcon("👥", Color.RoyalBlue);
            button2.Image = GetIcon("🏷️", Color.Chocolate);
            button1.Image = GetIcon("🔢", Color.DarkSlateGray);
            button3.Image = GetIcon("💸", Color.SeaGreen);
            button4.Image = GetIcon("🗣️", Color.DarkSlateBlue);
            button5.Image = GetIcon("📏", Color.DimGray);
            button6.Image = GetIcon("📖", Color.SaddleBrown);
            button7.Image = GetIcon("🏢", Color.SteelBlue);
            button11.Image = GetIcon("⚖️", Color.Sienna);
            button9.Image = GetIcon("🌍", Color.DodgerBlue);
            button10.Image = GetIcon("📏", Color.SlateGray);
            button12.Image = GetIcon("📁", Color.Goldenrod);

            // Colonna 2
            button36.Image = GetIcon("🛠️", Color.Gray);
            button15.Image = GetIcon("📄", Color.LightSlateGray);
            button16.Image = GetIcon("🔄", Color.CornflowerBlue);
            button17.Image = GetIcon("📑", Color.SlateGray);
            button34.Image = GetIcon("🧾", Color.DarkOliveGreen);
            button18.Image = GetIcon("💳", Color.DarkSlateBlue);
            button19.Image = GetIcon("📢", Color.OrangeRed);
            button21.Image = GetIcon("⚖️", Color.Sienna);
            button22.Image = GetIcon("⚖️", Color.Gray);
            button23.Image = GetIcon("🚚", Color.ForestGreen);
            button25.Image = GetIcon("🚛", Color.DarkSlateGray);
            button37.Image = GetIcon("🏪", Color.DarkSlateBlue);
            button13.Image = GetIcon("📋", Color.SlateGray);

            // Colonna 3
            button27.Image = GetIcon("👤", Color.DarkBlue);
            button29.Image = GetIcon("💳", Color.Indigo);
            button46.Image = GetIcon("🗺️", Color.DarkCyan);
            button30.Image = GetIcon("🎁", Color.Firebrick);
            button31.Image = GetIcon("📲", Color.CadetBlue);
            button14.Image = GetIcon("🏷️", Color.Chocolate);
            button41.Image = GetIcon("🏷️", Color.DarkOrange);
            button26.Image = GetIcon("🎨", Color.MediumVioletRed);
            button39.Image = GetIcon("📝", Color.DarkGoldenrod);
            button40.Image = GetIcon("📦", Color.Brown);
            button20.Image = GetIcon("🏧", Color.Navy);
            button42.Image = GetIcon("🏭", Color.SlateGray);
            button45.Image = GetIcon("👤", Color.Navy);

            // Colonna 4
            button24.Image = GetIcon("📥", Color.Goldenrod);
            button8.Image = GetIcon("📦", Color.Peru);
            button28.Image = GetIcon("🎫", Color.DarkMagenta);
            button33.Image = GetIcon("🏧", Color.MidnightBlue);
            button35.Image = GetIcon("📱", Color.Teal);
            button38.Image = GetIcon("🏠", Color.DarkCyan);
            button32.Image = GetIcon("💰", Color.DarkGreen);
            button44.Image = GetIcon("🛣️", Color.DarkSlateBlue);
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
                if (panel1 != null)
                {
                    panel1.BackColor = Color.FromArgb(248, 250, 252);
                }

                // Colonna 1: Anagrafiche, Configurazione e Tabelle Base
                clsUiIcons.StyleUtilityButtonWithImage(btnConf, btnConf.Image,
                    Color.FromArgb(238, 242, 255), Color.FromArgb(199, 210, 254), Color.FromArgb(129, 140, 248),
                    Color.FromArgb(49, 46, 129), Color.FromArgb(79, 70, 229));

                clsUiIcons.StyleUtilityButtonWithImage(button43, button43.Image,
                    Color.FromArgb(239, 246, 255), Color.FromArgb(191, 219, 254), Color.FromArgb(96, 165, 250),
                    Color.FromArgb(30, 58, 138), Color.FromArgb(29, 78, 216));

                clsUiIcons.StyleUtilityButtonWithImage(button2, button2.Image,
                    Color.FromArgb(255, 247, 237), Color.FromArgb(254, 215, 170), Color.FromArgb(251, 146, 60),
                    Color.FromArgb(124, 45, 18), Color.FromArgb(234, 88, 12));

                clsUiIcons.StyleUtilityButtonWithImage(button1, button1.Image,
                    Color.FromArgb(240, 249, 255), Color.FromArgb(186, 230, 253), Color.FromArgb(56, 189, 248),
                    Color.FromArgb(12, 74, 110), Color.FromArgb(2, 132, 199));

                clsUiIcons.StyleUtilityButtonWithImage(button3, button3.Image,
                    Color.FromArgb(255, 253, 245), Color.FromArgb(254, 230, 138), Color.FromArgb(251, 191, 36),
                    Color.FromArgb(120, 53, 15), Color.FromArgb(217, 119, 6));

                clsUiIcons.StyleUtilityButtonWithImage(button4, button4.Image,
                    Color.FromArgb(238, 242, 255), Color.FromArgb(199, 210, 254), Color.FromArgb(129, 140, 248),
                    Color.FromArgb(49, 46, 129), Color.FromArgb(79, 70, 229));

                clsUiIcons.StyleUtilityButtonWithImage(button5, button5.Image,
                    Color.FromArgb(241, 245, 249), Color.FromArgb(203, 213, 225), Color.FromArgb(148, 163, 184),
                    Color.FromArgb(30, 41, 59), Color.FromArgb(71, 85, 105));

                clsUiIcons.StyleUtilityButtonWithImage(button6, button6.Image,
                    Color.FromArgb(241, 245, 249), Color.FromArgb(203, 213, 225), Color.FromArgb(148, 163, 184),
                    Color.FromArgb(30, 41, 59), Color.FromArgb(71, 85, 105));

                clsUiIcons.StyleUtilityButtonWithImage(button7, button7.Image,
                    Color.FromArgb(236, 254, 255), Color.FromArgb(165, 243, 252), Color.FromArgb(56, 189, 248),
                    Color.FromArgb(22, 78, 99), Color.FromArgb(2, 132, 199));

                clsUiIcons.StyleUtilityButtonWithImage(button11, button11.Image,
                    Color.FromArgb(236, 254, 255), Color.FromArgb(165, 243, 252), Color.FromArgb(56, 189, 248),
                    Color.FromArgb(22, 78, 99), Color.FromArgb(2, 132, 199));

                clsUiIcons.StyleUtilityButtonWithImage(button9, button9.Image,
                    Color.FromArgb(240, 249, 255), Color.FromArgb(186, 230, 253), Color.FromArgb(56, 189, 248),
                    Color.FromArgb(12, 74, 110), Color.FromArgb(2, 132, 199));

                clsUiIcons.StyleUtilityButtonWithImage(button10, button10.Image,
                    Color.FromArgb(241, 245, 249), Color.FromArgb(203, 213, 225), Color.FromArgb(148, 163, 184),
                    Color.FromArgb(30, 41, 59), Color.FromArgb(71, 85, 105));

                clsUiIcons.StyleUtilityButtonWithImage(button12, button12.Image,
                    Color.FromArgb(254, 252, 232), Color.FromArgb(254, 240, 138), Color.FromArgb(250, 204, 21),
                    Color.FromArgb(113, 63, 18), Color.FromArgb(202, 138, 4));

                // Colonna 2: Documenti, Causali, Pagamenti, Negozi, Listini
                clsUiIcons.StyleUtilityButtonWithImage(button36, button36.Image,
                    Color.FromArgb(248, 250, 252), Color.FromArgb(203, 213, 225), Color.FromArgb(148, 163, 184),
                    Color.FromArgb(15, 23, 42), Color.FromArgb(51, 65, 85));

                clsUiIcons.StyleUtilityButtonWithImage(button15, button15.Image,
                    Color.FromArgb(255, 253, 245), Color.FromArgb(254, 230, 138), Color.FromArgb(251, 191, 36),
                    Color.FromArgb(120, 53, 15), Color.FromArgb(217, 119, 6));

                clsUiIcons.StyleUtilityButtonWithImage(button16, button16.Image,
                    Color.FromArgb(239, 246, 255), Color.FromArgb(191, 219, 254), Color.FromArgb(96, 165, 250),
                    Color.FromArgb(30, 58, 138), Color.FromArgb(29, 78, 216));

                clsUiIcons.StyleUtilityButtonWithImage(button17, button17.Image,
                    Color.FromArgb(254, 252, 232), Color.FromArgb(254, 240, 138), Color.FromArgb(250, 204, 21),
                    Color.FromArgb(113, 63, 18), Color.FromArgb(202, 138, 4));

                clsUiIcons.StyleUtilityButtonWithImage(button34, button34.Image,
                    Color.FromArgb(236, 253, 245), Color.FromArgb(167, 243, 208), Color.FromArgb(52, 211, 153),
                    Color.FromArgb(6, 78, 59), Color.FromArgb(5, 150, 105));

                clsUiIcons.StyleUtilityButtonWithImage(button18, button18.Image,
                    Color.FromArgb(238, 242, 255), Color.FromArgb(199, 210, 254), Color.FromArgb(129, 140, 248),
                    Color.FromArgb(49, 46, 129), Color.FromArgb(79, 70, 229));

                clsUiIcons.StyleUtilityButtonWithImage(button19, button19.Image,
                    Color.FromArgb(255, 247, 237), Color.FromArgb(254, 215, 170), Color.FromArgb(251, 146, 60),
                    Color.FromArgb(124, 45, 18), Color.FromArgb(234, 88, 12));

                clsUiIcons.StyleUtilityButtonWithImage(button21, button21.Image,
                    Color.FromArgb(254, 252, 232), Color.FromArgb(254, 240, 138), Color.FromArgb(250, 204, 21),
                    Color.FromArgb(113, 63, 18), Color.FromArgb(202, 138, 4));

                clsUiIcons.StyleUtilityButtonWithImage(button22, button22.Image,
                    Color.FromArgb(241, 245, 249), Color.FromArgb(203, 213, 225), Color.FromArgb(148, 163, 184),
                    Color.FromArgb(30, 41, 59), Color.FromArgb(71, 85, 105));

                clsUiIcons.StyleUtilityButtonWithImage(button23, button23.Image,
                    Color.FromArgb(240, 253, 250), Color.FromArgb(153, 246, 228), Color.FromArgb(45, 212, 191),
                    Color.FromArgb(19, 78, 74), Color.FromArgb(13, 148, 136));

                clsUiIcons.StyleUtilityButtonWithImage(button25, button25.Image,
                    Color.FromArgb(240, 253, 244), Color.FromArgb(187, 247, 208), Color.FromArgb(74, 222, 128),
                    Color.FromArgb(20, 83, 45), Color.FromArgb(22, 163, 74));

                clsUiIcons.StyleUtilityButtonWithImage(button37, button37.Image,
                    Color.FromArgb(239, 246, 255), Color.FromArgb(191, 219, 254), Color.FromArgb(96, 165, 250),
                    Color.FromArgb(30, 58, 138), Color.FromArgb(29, 78, 216));

                clsUiIcons.StyleUtilityButtonWithImage(button13, button13.Image,
                    Color.FromArgb(238, 242, 255), Color.FromArgb(199, 210, 254), Color.FromArgb(129, 140, 248),
                    Color.FromArgb(49, 46, 129), Color.FromArgb(79, 70, 229));

                // Colonna 3: Tessere, Fidelity, Offerte, Note, POS, Magazzini
                clsUiIcons.StyleUtilityButtonWithImage(button27, button27.Image,
                    Color.FromArgb(239, 246, 255), Color.FromArgb(191, 219, 254), Color.FromArgb(96, 165, 250),
                    Color.FromArgb(30, 58, 138), Color.FromArgb(29, 78, 216));

                clsUiIcons.StyleUtilityButtonWithImage(button29, button29.Image,
                    Color.FromArgb(245, 243, 255), Color.FromArgb(221, 214, 254), Color.FromArgb(167, 139, 250),
                    Color.FromArgb(76, 29, 149), Color.FromArgb(124, 58, 237));

                clsUiIcons.StyleUtilityButtonWithImage(button46, button46.Image,
                    Color.FromArgb(240, 253, 250), Color.FromArgb(153, 246, 228), Color.FromArgb(45, 212, 191),
                    Color.FromArgb(19, 78, 74), Color.FromArgb(13, 148, 136));

                clsUiIcons.StyleUtilityButtonWithImage(button30, button30.Image,
                    Color.FromArgb(245, 243, 255), Color.FromArgb(221, 214, 254), Color.FromArgb(167, 139, 250),
                    Color.FromArgb(76, 29, 149), Color.FromArgb(124, 58, 237));

                clsUiIcons.StyleUtilityButtonWithImage(button31, button31.Image,
                    Color.FromArgb(236, 254, 255), Color.FromArgb(165, 243, 252), Color.FromArgb(56, 189, 248),
                    Color.FromArgb(22, 78, 99), Color.FromArgb(2, 132, 199));

                clsUiIcons.StyleUtilityButtonWithImage(button14, button14.Image,
                    Color.FromArgb(255, 247, 237), Color.FromArgb(254, 215, 170), Color.FromArgb(251, 146, 60),
                    Color.FromArgb(124, 45, 18), Color.FromArgb(234, 88, 12));

                clsUiIcons.StyleUtilityButtonWithImage(button41, button41.Image,
                    Color.FromArgb(255, 247, 237), Color.FromArgb(254, 215, 170), Color.FromArgb(251, 146, 60),
                    Color.FromArgb(124, 45, 18), Color.FromArgb(234, 88, 12));

                clsUiIcons.StyleUtilityButtonWithImage(button26, button26.Image,
                    Color.FromArgb(255, 241, 242), Color.FromArgb(254, 205, 211), Color.FromArgb(251, 113, 133),
                    Color.FromArgb(136, 19, 55), Color.FromArgb(225, 29, 72));

                clsUiIcons.StyleUtilityButtonWithImage(button39, button39.Image,
                    Color.FromArgb(254, 252, 232), Color.FromArgb(254, 240, 138), Color.FromArgb(250, 204, 21),
                    Color.FromArgb(113, 63, 18), Color.FromArgb(202, 138, 4));

                clsUiIcons.StyleUtilityButtonWithImage(button40, button40.Image,
                    Color.FromArgb(255, 241, 242), Color.FromArgb(254, 205, 211), Color.FromArgb(251, 113, 133),
                    Color.FromArgb(136, 19, 55), Color.FromArgb(225, 29, 72));

                clsUiIcons.StyleUtilityButtonWithImage(button20, button20.Image,
                    Color.FromArgb(238, 242, 255), Color.FromArgb(199, 210, 254), Color.FromArgb(129, 140, 248),
                    Color.FromArgb(49, 46, 129), Color.FromArgb(79, 70, 229));

                clsUiIcons.StyleUtilityButtonWithImage(button42, button42.Image,
                    Color.FromArgb(248, 250, 252), Color.FromArgb(203, 213, 225), Color.FromArgb(148, 163, 184),
                    Color.FromArgb(15, 23, 42), Color.FromArgb(51, 65, 85));

                clsUiIcons.StyleUtilityButtonWithImage(button45, button45.Image,
                    Color.FromArgb(239, 246, 255), Color.FromArgb(191, 219, 254), Color.FromArgb(96, 165, 250),
                    Color.FromArgb(30, 58, 138), Color.FromArgb(29, 78, 216));

                // Colonna 4: Import ECR, Merceologie, Tessere, Divulgazioni, Pagamenti cassa, Tracciati
                clsUiIcons.StyleUtilityButtonWithImage(button24, button24.Image,
                    Color.FromArgb(254, 252, 232), Color.FromArgb(254, 240, 138), Color.FromArgb(250, 204, 21),
                    Color.FromArgb(113, 63, 18), Color.FromArgb(202, 138, 4));

                clsUiIcons.StyleUtilityButtonWithImage(button8, button8.Image,
                    Color.FromArgb(240, 253, 250), Color.FromArgb(153, 246, 228), Color.FromArgb(45, 212, 191),
                    Color.FromArgb(19, 78, 74), Color.FromArgb(13, 148, 136));

                clsUiIcons.StyleUtilityButtonWithImage(button28, button28.Image,
                    Color.FromArgb(245, 243, 255), Color.FromArgb(221, 214, 254), Color.FromArgb(167, 139, 250),
                    Color.FromArgb(76, 29, 149), Color.FromArgb(124, 58, 237));

                clsUiIcons.StyleUtilityButtonWithImage(button33, button33.Image,
                    Color.FromArgb(236, 254, 255), Color.FromArgb(165, 243, 252), Color.FromArgb(56, 189, 248),
                    Color.FromArgb(22, 78, 99), Color.FromArgb(2, 132, 199));

                clsUiIcons.StyleUtilityButtonWithImage(button35, button35.Image,
                    Color.FromArgb(236, 254, 255), Color.FromArgb(165, 243, 252), Color.FromArgb(56, 189, 248),
                    Color.FromArgb(22, 78, 99), Color.FromArgb(2, 132, 199));

                clsUiIcons.StyleUtilityButtonWithImage(button38, button38.Image,
                    Color.FromArgb(240, 253, 250), Color.FromArgb(153, 246, 228), Color.FromArgb(45, 212, 191),
                    Color.FromArgb(19, 78, 74), Color.FromArgb(13, 148, 136));

                clsUiIcons.StyleUtilityButtonWithImage(button32, button32.Image,
                    Color.FromArgb(236, 253, 245), Color.FromArgb(167, 243, 208), Color.FromArgb(52, 211, 153),
                    Color.FromArgb(6, 78, 59), Color.FromArgb(5, 150, 105));

                clsUiIcons.StyleUtilityButtonWithImage(button44, button44.Image,
                    Color.FromArgb(238, 242, 255), Color.FromArgb(199, 210, 254), Color.FromArgb(129, 140, 248),
                    Color.FromArgb(49, 46, 129), Color.FromArgb(79, 70, 229));
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("frmMnuUtilita.ApplyModernUi", ex.Message);
            }
        }

        private Image GetIcon(string unicodeChar, Color fallbackColor)
        {
            Bitmap bmp = new Bitmap(24, 24);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
                using (Font f = new Font("Segoe UI Emoji", 12, FontStyle.Regular))
                using (SolidBrush b = new SolidBrush(fallbackColor))
                {
                    StringFormat fmt = new StringFormat();
                    fmt.Alignment = StringAlignment.Center;
                    fmt.LineAlignment = StringAlignment.Center;
                    g.DrawString(unicodeChar, f, b, new RectangleF(0, 0, 24, 24), fmt);
                }
            }
            return bmp;
        }

        private void frmMnuUtilita_KeyDown(object sender, KeyEventArgs e)
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
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            frmGesTabelle f = new frmGesTabelle();
            f._strTab = "TabNumeratori";
            f._strTabDes = "Numeratori";
            f._bolCodAlf = false;
            f._intCodLen = 3;
            f._intDesLen = 50;
            f._bolColAnn = true;
            f.ShowDialog();
        }

        //private void button2_Click(object sender, EventArgs e)
        //{
        //    frmGesTabelle f = new frmGesTabelle();
        //    f._strTab = "TabParametri";
        //    f._strTabDes = "Parametri";
        //    f._bolCodAlf = false;
        //    f._intCodLen = 3;
        //    f._intDesLen = 50;
        //    f._bolColAnn = true;
        //    f.ShowDialog();
        //}

        private void button3_Click(object sender, EventArgs e)
        {
            frmGesTabelle f = new frmGesTabelle();
            f._strTab = "TabIva";
            f._strTabDes = "IVA";
            f._bolCodAlf = false;
            f._intCodLen = 3;
            f._intDesLen = 20;
            f._bolColAnn = true;
            f.ShowDialog();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            frmGesTabelle f = new frmGesTabelle();
            f._strTab = "TabStato";
            f._strTabDes = "Stato voce";
            f._bolCodAlf = true;
            f._intCodLen = 1;
            f._intDesLen = 20;
            f._bolColAnn = true;
            f.ShowDialog();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            frmGesTabelle f = new frmGesTabelle();
            f._strTab = "TabUmi";
            f._strTabDes = "Unità di misura";
            f._bolCodAlf = true;
            f._intCodLen = 2;
            f._intDesLen = 10;
            f._bolColAnn = true;
            f.ShowDialog();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            frmGesTabelle f = new frmGesTabelle();
            f._strTab = "TabUmiEtichetta";
            f._strTabDes = "Unità di misura in chiaro";
            f._bolCodAlf = true;
            f._intCodLen = 2;
            f._intDesLen = 10;
            f._bolColAnn = true;
            f.ShowDialog();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            frmGesTabelle f = new frmGesTabelle();
            f._strTab = "TabReparti";
            f._strTabDes = "Reparti";
            f._bolCodAlf = true;
            f._intCodLen = 3;
            f._intDesLen = 20;
            f._bolColAnn = true;
            f.ShowDialog();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            new frmGesTabEcr().ShowDialog();
        }
 
        private void button9_Click(object sender, EventArgs e)
        {
            frmGesTabelle f = new frmGesTabelle();
            f._strTab = "TabOrigine";
            f._strTabDes = "Origine";
            f._bolCodAlf = true;
            f._intCodLen = 3;
            f._intDesLen = 20;
            f._bolColAnn = true;
            f.ShowDialog();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            frmGesTabelle f = new frmGesTabelle();
            f._strTab = "TabCalibro";
            f._strTabDes = "Calibro";
            f._bolCodAlf = true;
            f._intCodLen = 3;
            f._intDesLen = 20;
            f._bolColAnn = true;
            f.ShowDialog();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            frmGesTabelle f = new frmGesTabelle();
            f._strTab = "TabRepBilance";
            f._strTabDes = "Reparti bilancia";
            f._bolCodAlf = true;
            f._intCodLen = 1;
            f._intDesLen = 20;
            f._bolColAnn = true;
            f.ShowDialog();
        }

        private void button12_Click(object sender, EventArgs e)
        {
            frmGesTabelle f = new frmGesTabelle();
            f._strTab = "TabCategoria";
            f._strTabDes = "Categoria";
            f._bolCodAlf = true;
            f._intCodLen = 1;
            f._intDesLen = 20;
            f._bolColAnn = true;
            f.ShowDialog();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            frmGesTabelle f = new frmGesTabelle();
            f._strTab = "TabListini";
            f._strTabDes = "Listini";
            f._bolCodAlf = true;
            f._intCodLen = 3;
            f._intDesLen = 20;
            f._bolColAnn = true;
            f.ShowDialog();
        }

        private void button14_Click(object sender, EventArgs e)
        {
            frmGesTabelle f = new frmGesTabelle();
            f._strTab = "TabOfferte";
            f._strTabDes = "Tipo offerte";
            f._bolCodAlf = true;
            f._intCodLen = 3;
            f._intDesLen = 20;
            f._bolColAnn = true;
            f.ShowDialog();
        }

        private void button15_Click(object sender, EventArgs e)
        {
            frmGesTabelle f = new frmGesTabelle();
            f._strTab = "TabDocumenti";
            f._strTabDes = "Documenti";
            f._bolCodAlf = true;
            f._intCodLen = 1;
            f._intDesLen = 20;
            f._bolColAnn = true;
            f.ShowDialog();
        }

        private void button16_Click(object sender, EventArgs e)
        {
            frmGesTabelle f = new frmGesTabelle();
            f._strTab = "TabMovCausali";
            f._strTabDes = "Causali movimenti";
            f._bolCodAlf = false;
            f._intCodLen = 3;
            f._intDesLen = 20;
            f._bolColAnn = true;
            f.ShowDialog();
        }

        private void button17_Click(object sender, EventArgs e)
        {
            frmGesTabelle f = new frmGesTabelle();
            f._strTab = "TabDocTipo";
            f._strTabDes = "Tipo documenti";
            f._bolCodAlf = true;
            f._intCodLen = 2;
            f._intDesLen = 20;
            f._bolColAnn = true;
            f.ShowDialog();
        }

        private void button18_Click(object sender, EventArgs e)
        {
            frmGesTabelle f = new frmGesTabelle();
            f._strTab = "TabPagamenti";
            f._strTabDes = "Tipo pagamenti";
            f._bolCodAlf = true;
            f._intCodLen = 3;
            f._intDesLen = 30;
            f._bolColAnn = true;
            f.ShowDialog();
        }

        private void button19_Click(object sender, EventArgs e)
        {
            frmGesTabelle f = new frmGesTabelle();
            f._strTab = "TabStatoDivulgazioni";
            f._strTabDes = "Stato divulgazioni";
            f._bolCodAlf = true;
            f._intCodLen = 1;
            f._intDesLen = 20;
            f._bolColAnn = true;
            f.ShowDialog();
        }

        private void btnConf_Click(object sender, EventArgs e)
        {
            frmUser f = new frmUser();
            f._strConSql = _clsFun.ConSql("");
            f._strTip = "CNF";
            int num = (int)f.ShowDialog();
            if (f._bolSta)
                new frmUtyConfig().ShowDialog();
        }

        private void button20_Click(object sender, EventArgs e)
        {
            frmGesTabelle f = new frmGesTabelle();
            f._strTab = "TabPos";
            f._strTabDes = "POS";
            f._bolCodAlf = true;
            f._intCodLen = 3;
            f._intDesLen = 20;
            f._bolColAnn = true;
            f.ShowDialog();
        }

        private void button21_Click(object sender, EventArgs e)
        {
            frmGesTabelle f = new frmGesTabelle();
            f._strTab = "TabBilance";
            f._strTabDes = "Bilance";
            f._bolCodAlf = true;
            f._intCodLen = 3;
            f._intDesLen = 20;
            f._bolColAnn = true;
            f.ShowDialog();
        }

        private void button22_Click(object sender, EventArgs e)
        {
            frmGesTabelle f = new frmGesTabelle();
            f._strTab = "TabTipoGrammatura";
            f._strTabDes = "T.grammatura";
            f._bolCodAlf = true;
            f._intCodLen = 2;
            f._intDesLen = 20;
            f._bolColAnn = true;
            f.ShowDialog();
        }

        private void button23_Click(object sender, EventArgs e)
        {
            string s = Path.GetDirectoryName(_clsDef.FILEINI) + "\\ApForTracciati.mdb";
            if (!File.Exists(s))
                MessageBox.Show("File tracciati mancante!", "CONTROLLO PARAMETRI", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                frmGesTabelle f = new frmGesTabelle();
                f._strTab = "TabForImport";
                f._strTabDes = "Divulgazioni fornitori";
                f._bolCodAlf = true;
                f._intCodLen = 5;
                f._intDesLen = 50;
                f._bolColAnn = true;
                f.ShowDialog();
            }
        }

        private void button24_Click(object sender, EventArgs e)
        {
            new frmUtyImpEcr().ShowDialog();
        }

        private void button25_Click(object sender, EventArgs e)
        {
            frmGesTabelle f = new frmGesTabelle();
            f._strTab = "TabCauTrasporto";
            f._strTabDes = "Causali trasporto";
            f._bolCodAlf = true;
            f._intCodLen = 3;
            f._intDesLen = 20;
            f._bolColAnn = true;
            f.ShowDialog();
        }

        private void button26_Click(object sender, EventArgs e)
        {
            frmGesTabelle f = new frmGesTabelle();
            f._strTab = "TabAspettoBeni";
            f._strTabDes = "Aspetto esteriore beni";
            f._bolCodAlf = true;
            f._intCodLen = 3;
            f._intDesLen = 20;
            f._bolColAnn = true;
            f.ShowDialog();
        }

        private void button27_Click(object sender, EventArgs e)
        {
            frmGesTabelle f = new frmGesTabelle();
            f._strTab = "TabProfessioni";
            f._strTabDes = "Impieghi";
            f._bolCodAlf = false;
            f._intCodLen = 3;
            f._intDesLen = 20;
            f._bolColAnn = true;
            f.ShowDialog();
        }

        private void button28_Click(object sender, EventArgs e)
        {
            new frmUtyTessere().ShowDialog();
        }

        private void button29_Click(object sender, EventArgs e)
        {
            frmGesTabelle f = new frmGesTabelle();
            f._strTab = "TabFidGruppi";
            f._strTabDes = "Gruppi tessere";
            f._bolCodAlf = false;
            f._intCodLen = 3;
            f._intDesLen = 20;
            f._bolColAnn = true;
            f.ShowDialog();
        }

        private void button30_Click(object sender, EventArgs e)
        {
            frmGesTabelle f = new frmGesTabelle();
            f._strTab = "TabFidCampagne";
            f._strTabDes = "Campagne fidelity";
            f._bolCodAlf = false;
            f._intCodLen = 3;
            f._intDesLen = 20;
            f._bolColAnn = true;
            f.ShowDialog();
        }

        private void button31_Click(object sender, EventArgs e)
        {
            frmGesTabelle f = new frmGesTabelle();
            f._strTab = "TabTerm";
            f._strTabDes = "Terminalini";
            f._bolCodAlf = false;
            f._intCodLen = 3;
            f._intDesLen = 20;
            f._bolColAnn = true;
            f.ShowDialog();
        }

        private void button32_Click(object sender, EventArgs e)
        {
            frmGesTabelle f = new frmGesTabelle();
            f._strTab = "TabCauCassa";
            f._strTabDes = "Pagamenti cassa";
            f._bolCodAlf = false;
            f._intCodLen = 3;
            f._intDesLen = 20;
            f._bolColAnn = true;
            f.ShowDialog();
        }

        private void button33_Click(object sender, EventArgs e)
        {
            DataTable tCnf = new clsQuery().ConfSeek("", "POS");
            if (tCnf.Rows.Count > 0 && (string)tCnf.Rows[0]["cnf_pos"] == "05")
                new frmUtyDivPos().ShowDialog();
            else
                MessageBox.Show("Funzione attiva solo per casse ApShop Touch!");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            frmGesTabelle f = new frmGesTabelle();
            f._strTab = "TabEtichette";
            f._strTabDes = "Etichette";
            f._bolCodAlf = true;
            f._intCodLen = 3;
            f._intDesLen = 20;
            f._bolColAnn = true;
            f.ShowDialog();
        }

        private void button34_Click(object sender, EventArgs e)
        {
            frmGesTabelle f = new frmGesTabelle();
            f._strTab = "TabDocTpd";
            f._strTabDes = "Tipo documento fattura";
            f._bolCodAlf = true;
            f._intCodLen = 2;
            f._intDesLen = 20;
            f._bolColAnn = true;
            f.ShowDialog();
        }

        private void button37_Click(object sender, EventArgs e)
        {
            frmGesTabelle f = new frmGesTabelle();
            f._strTab = "TabNegozi";
            f._strTabDes = "Negozi";
            f._bolCodAlf = true;
            f._intCodLen = 3;
            f._intDesLen = 20;
            f._bolColAnn = true;
            f.ShowDialog();
        }

        private void button35_Click(object sender, EventArgs e)
        {
            string s = "SELECT tab_cod FROM TabTerm WHERE tab_ann=0 AND tab_cod='" + _clsDef.TERMTDIV + "'";
            DataTable t = _clsFun.FillTabSql("TabTerm", s, true, _strConSql);
            if (t.Rows.Count > 0 )
                new frmUtyDivTerm().ShowDialog();
            else
                MessageBox.Show("Funzione attiva solo per divulgazione Terminalini TMPDIV attiva!");
        }

        private void button36_Click(object sender, EventArgs e)
        {
            new frmMnuUtility().ShowDialog();
        }

        private void button38_Click(object sender, EventArgs e)
        {
            new frmUtyDivNegozi().ShowDialog();
        }

        private void button39_Click(object sender, EventArgs e)
        {
            frmGesTabelle f = new frmGesTabelle();
            f._strTab = "TabNote";
            f._strTabDes = "Note";
            f._bolCodAlf = true;
            f._intCodLen = 2;
            f._intDesLen = 20;
            f._bolColAnn = true;
            f.ShowDialog();
        }

        private void button40_Click(object sender, EventArgs e)
        {
            frmGesTabelle f = new frmGesTabelle();
            f._strTab = "TabLottoTipo";
            f._strTabDes = "Lotto tipo";
            f._bolCodAlf = true;
            f._intCodLen = 1;
            f._intDesLen = 20;
            f._bolColAnn = true;
            f.ShowDialog();
        }

        private void button42_Click(object sender, EventArgs e)
        {
            frmGesTabelle f = new frmGesTabelle();
            f._strTab = "TabMagazzini";
            f._strTabDes = "Magazzini";
            f._bolCodAlf = true;
            f._intCodLen = 2;
            f._intDesLen = 20;
            f._bolColAnn = true;
            f.ShowDialog();
        }

        private void button45_Click(object sender, EventArgs e)
        {
            frmGesTabelle f = new frmGesTabelle();
            f._strTab = "TabClienteTipo";
            f._strTabDes = "Tipo cliente";
            f._bolCodAlf = true;
            f._intCodLen = 1;
            f._intDesLen = 20;
            f._bolColAnn = true;
            f.ShowDialog();
        }

        private void button41_Click(object sender, EventArgs e)
        {
            new frmAnaDivFornitoriOld().ShowDialog();
        }

        private void button43_Click(object sender, EventArgs e)
        {
            new frmAnaGestione().ShowDialog();
        }

        private void button44_Click(object sender, EventArgs e)
        {
            new frmAnaDivFornitori().ShowDialog();
        }

        private void button41_Click_1(object sender, EventArgs e)
        {
            frmGesTabelle f = new frmGesTabelle();
            f._strTab = "TabOffTransazione";
            f._strTabDes = "Tipo offerte transazione";
            f._bolCodAlf = true;
            f._intCodLen = 3;
            f._intDesLen = 50;
            f._bolColAnn = true;
            f.ShowDialog();
        }

        private void button46_Click(object sender, EventArgs e)
        {
            frmGesTabelle f = new frmGesTabelle();
            f._strTab = "TabProvince";
            f._strTabDes = "Province";
            f._bolCodAlf = true;
            f._intCodLen = 2;
            f._intDesLen = 20;
            f._bolColAnn = true;
            f.ShowDialog();
        }

        //private void button41_Click_1(object sender, EventArgs e)
        //{
        //    string s = "";
        //    s = "\\\\192.168.16.9\\PuntiVendita\\Sede\\DivNegozi\\";
        //    //s = "\\\\192.168.178.37\\scambio\\apsistemi\\";
        //    if (DirectoryExists(s))
        //        MessageBox.Show("Passo");
        //    else
        //        MessageBox.Show("NON Passo");
        //}

        //private bool DirectoryExists(string directory)
        //{
        //    Func<bool> func = () => Directory.Exists(directory);
        //    Task<bool> task = new Task<bool>(func);
        //    task.Start();
        //    if (task.Wait(500))
        //    {
        //        return task.Result;
        //    }
        //    else
        //    {
        //        // Didn't get an answer back in time be pessimistic and assume it didn't exist
        //        return false;
        //    }
        //    //Boolean b = Directory.Exists(directory);
        //    //return b;
        //}

   }
}
