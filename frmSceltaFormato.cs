using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;

namespace APOffice
{
    public partial class frmSceltaFormato : Form
    {
        public DataTable _tabFormati;
        public string _strScelta = "";

        private ListBox lstFormati;
        private Button btnOk;
        private Button btnAnnulla;

        public frmSceltaFormato()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.lstFormati = new System.Windows.Forms.ListBox();
            this.btnOk = new System.Windows.Forms.Button();
            this.btnAnnulla = new System.Windows.Forms.Button();
            this.SuspendLayout();
            
            // lstFormati
            this.lstFormati.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lstFormati.FormattingEnabled = true;
            this.lstFormati.ItemHeight = 16;
            this.lstFormati.Location = new System.Drawing.Point(12, 12);
            this.lstFormati.Name = "lstFormati";
            this.lstFormati.Size = new System.Drawing.Size(260, 212);
            this.lstFormati.TabIndex = 0;
            this.lstFormati.DoubleClick += new System.EventHandler(this.btnOk_Click);

            // btnOk
            this.btnOk.Location = new System.Drawing.Point(12, 230);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(90, 30);
            this.btnOk.TabIndex = 1;
            this.btnOk.Text = "Conferma";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            
            // btnAnnulla
            this.btnAnnulla.Location = new System.Drawing.Point(182, 230);
            this.btnAnnulla.Name = "btnAnnulla";
            this.btnAnnulla.Size = new System.Drawing.Size(90, 30);
            this.btnAnnulla.TabIndex = 2;
            this.btnAnnulla.Text = "Annulla";
            this.btnAnnulla.UseVisualStyleBackColor = true;
            this.btnAnnulla.Click += new System.EventHandler(this.btnAnnulla_Click);

            // frmSceltaFormato
            this.AcceptButton = this.btnOk;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnAnnulla;
            this.ClientSize = new System.Drawing.Size(284, 271);
            this.Controls.Add(this.btnAnnulla);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.lstFormati);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmSceltaFormato";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Scelta Formato Etichetta";
            this.Load += new System.EventHandler(this.frmSceltaFormato_Load);
            this.ResumeLayout(false);
        }

        private void frmSceltaFormato_Load(object sender, EventArgs e)
        {
            if (_tabFormati != null && _tabFormati.Rows.Count > 0)
            {
                // Create a bindable list of displays
                foreach (DataRow r in _tabFormati.Rows)
                {
                    string cod = r["cod"].ToString();
                    string des = r["des"].ToString();
                    lstFormati.Items.Add(cod + " - " + des);
                }
                lstFormati.SelectedIndex = 0;
            }
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            if (lstFormati.SelectedIndex >= 0)
            {
                string s = lstFormati.SelectedItem.ToString();
                // Extact code before space
                string[] parts = s.Split(' ');
                if (parts.Length > 0)
                    _strScelta = parts[0].Trim();
            }
            this.Close();
        }

        private void btnAnnulla_Click(object sender, EventArgs e)
        {
            _strScelta = "";
            this.Close();
        }
    }
}
