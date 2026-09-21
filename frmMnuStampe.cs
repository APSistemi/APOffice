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
    public partial class frmMnuStampe : Form
    {
        public frmMnuStampe()
        {
            InitializeComponent(); 
            new clsGesGraph().SetGraph(this, 0);
        }

        private void frmMnuStampe_Load(object sender, EventArgs e)
        {
            clsUiIcons.RestoreFormBounds(this);
            ApplyModernUi();
        }

        private void ApplyModernUi()
        {
            try
            {
                this.BackColor = Color.FromArgb(243, 244, 246);
                if (panel1 != null)
                {
                    panel1.BackColor = Color.FromArgb(248, 250, 252);
                }

                if (menuStrip1 != null)
                {
                    menuStrip1.Renderer = clsUiIcons.GetModernMenuRenderer();
                    if (esciToolStripMenuItem != null)
                        esciToolStripMenuItem.Image = clsUiIcons.GetIcon("exit", 16);
                }

                // 1. Stampa PLU bilancia (Sky Blue)
                clsUiIcons.StyleStatButton(button1, "Stampa PLU bilancia", "scale", 16,
                    Color.FromArgb(240, 249, 255), Color.FromArgb(186, 230, 253), Color.FromArgb(56, 189, 248),
                    Color.FromArgb(12, 74, 110), Color.FromArgb(2, 132, 199));

                // 2. Stampa articoli per fornitore (Emerald Green)
                clsUiIcons.StyleStatButton(button8, "Stampa articoli per fornitore", "truck", 16,
                    Color.FromArgb(240, 253, 244), Color.FromArgb(187, 247, 208), Color.FromArgb(74, 222, 128),
                    Color.FromArgb(20, 83, 45), Color.FromArgb(22, 163, 74));

                // 3. Stampa articoli per merceologia (Teal)
                clsUiIcons.StyleStatButton(button9, "Stampa articoli per merceologia", "categories", 16,
                    Color.FromArgb(240, 253, 250), Color.FromArgb(153, 246, 228), Color.FromArgb(45, 212, 191),
                    Color.FromArgb(19, 78, 74), Color.FromArgb(13, 148, 136));

                // 4. Stampa articoli per scadenza (Amber)
                clsUiIcons.StyleStatButton(button10, "Stampa articoli per scadenza", "calendar_week", 16,
                    Color.FromArgb(254, 252, 232), Color.FromArgb(254, 240, 138), Color.FromArgb(250, 204, 21),
                    Color.FromArgb(113, 63, 18), Color.FromArgb(202, 138, 4));

                // 5. Stampa giacenza (Cyan Blue)
                clsUiIcons.StyleStatButton(button7, "Stampa giacenza", "box", 16,
                    Color.FromArgb(239, 246, 255), Color.FromArgb(191, 219, 254), Color.FromArgb(96, 165, 250),
                    Color.FromArgb(30, 58, 138), Color.FromArgb(29, 78, 216));

                // 6. Disattivazione articoli (Soft Red)
                clsUiIcons.StyleStatButton(button2, "Disattivazione articoli", "trash", 16,
                    Color.FromArgb(254, 242, 242), Color.FromArgb(254, 202, 202), Color.FromArgb(248, 113, 113),
                    Color.FromArgb(127, 29, 29), Color.FromArgb(220, 38, 38));

                // 7. Controllo articoli (Slate Gray)
                clsUiIcons.StyleStatButton(button40, "Controllo articoli", "check", 16,
                    Color.FromArgb(241, 245, 249), Color.FromArgb(203, 213, 225), Color.FromArgb(148, 163, 184),
                    Color.FromArgb(30, 41, 59), Color.FromArgb(71, 85, 105));

                // 8. Estrazioni articoli (Purple Violet)
                clsUiIcons.StyleStatButton(button3, "Estrazioni articoli", "extract", 16,
                    Color.FromArgb(245, 243, 255), Color.FromArgb(221, 214, 254), Color.FromArgb(167, 139, 250),
                    Color.FromArgb(76, 29, 149), Color.FromArgb(124, 58, 237));

                // 9. Calendario storico statistiche (Indigo)
                clsUiIcons.StyleStatButton(button4, "Calendario storico statistiche", "statistics", 16,
                    Color.FromArgb(238, 242, 255), Color.FromArgb(199, 210, 254), Color.FromArgb(129, 140, 248),
                    Color.FromArgb(49, 46, 129), Color.FromArgb(79, 70, 229));

                // 10. Importazioni da APPHONE (Orange)
                clsUiIcons.StyleStatButton(button5, "Importazioni da APPHONE", "import", 16,
                    Color.FromArgb(255, 247, 237), Color.FromArgb(254, 215, 170), Color.FromArgb(251, 146, 60),
                    Color.FromArgb(124, 45, 18), Color.FromArgb(234, 88, 12));

                // 11. Controllo prezzi negozio (Teal Cyan)
                clsUiIcons.StyleStatButton(button6, "Controllo prezzi negozio", "discount", 16,
                    Color.FromArgb(236, 254, 255), Color.FromArgb(165, 243, 252), Color.FromArgb(56, 189, 248),
                    Color.FromArgb(22, 78, 99), Color.FromArgb(2, 132, 199));

                // 12. Genera Costi (Mint Emerald)
                clsUiIcons.StyleStatButton(button11, "Genera Costi", "refresh", 16,
                    Color.FromArgb(236, 253, 245), Color.FromArgb(167, 243, 208), Color.FromArgb(52, 211, 153),
                    Color.FromArgb(6, 78, 59), Color.FromArgb(5, 150, 105));
            }
            catch (Exception ex)
            {
                new clsFuncs().ErrorLog("frmMnuStampe.ApplyModernUi", ex.Message);
            }
        }

        private void frmMnuStampe_KeyDown(object sender, KeyEventArgs e)
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

        private void frmMnuStampe_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                clsUiIcons.SaveFormBounds(this);
            }
            catch { }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            new frmPrnPlu().ShowDialog();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            new frmPrnForArticoli().ShowDialog();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            new frmPrnArtMerceologie().ShowDialog();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            new frmPrnArtScadenza().ShowDialog();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            new frmUtyPrnGiacenza().ShowDialog();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            new frmUtyCtrlArts().ShowDialog();
        }

        private void button40_Click(object sender, EventArgs e)
        {
            new frmUtyAggArticoli().ShowDialog();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            new frmUtyArtEstrazioni().ShowDialog();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            new frmUtyStaStoCalendario().ShowDialog();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            new frmUtyDivApPhone().ShowDialog();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            new frmUtyPrzNegozio1().ShowDialog();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            new frmUtyGenCosti().ShowDialog();
        }
    }
}
