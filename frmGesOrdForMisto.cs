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
    public partial class frmGesOrdForMisto : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        public DataTable _tabOrd = new DataTable();

        private string _strConSql = "";

        public frmGesOrdForMisto()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0); ;
            _strConSql = _clsFun.ConSql("");
        }

        private void frmGesOrdForMisto_Load(object sender, EventArgs e)
        {
            SetDgv1();
            FillTabs();
            FillDati();
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }
        private void frmGesOrdForMisto_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
        }
        private void Esci()
        {
            _tabOrd = (DataTable)dgv1.DataSource;
            this.Close();
        }

        private void SetDgv1()
        {
            DataTable tFor = _clsFun.FillTabSql("AnaFornitori", "SELECT for_cod, for_des FROM AnaFornitori", false, _strConSql);

            dgv1.AutoGenerateColumns = false;
            //dgv1.VirtualMode = true;
            //dgv1.Dock = DockStyle.Fill;
            dgv1.AllowUserToAddRows = false;
            dgv1.ReadOnly = true;
            dgv1.AllowUserToDeleteRows = false;
            //dgv1.DisplayedRowCount() = true;

            DataGridViewTextBoxColumn cTbc;
            //DataGridViewCheckBoxColumn cCbc;
            DataGridViewComboBoxColumn cCmb;

            cCmb = new DataGridViewComboBoxColumn();
            cCmb.DataPropertyName = "tmp_for";
            cCmb.Name = "Fornitore";
            cCmb.Width = 120;
            cCmb.DataSource = tFor;
            cCmb.ValueMember = "for_cod";
            cCmb.DisplayMember = "for_des";
            cCmb.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing;
            cCmb.ReadOnly = true;
            dgv1.Columns.Add((DataGridViewColumn)cCmb);

            //cTbc = new DataGridViewTextBoxColumn();
            //cTbc.DataPropertyName = "orf_nri";
            //cTbc.Name = "Riga";
            //cTbc.Width = 30;
            //cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            //cTbc.ReadOnly = true;
            //dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "orf_arf";
            cTbc.Name = "Cod. fornitore";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "orf_art";
            cTbc.Name = "Articolo";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "orf_ard";
            cTbc.Name = "Descrizione";
            cTbc.Width = 200;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 50;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "orf_umi";
            cTbc.Name = "UM";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.MaxInputLength = 50;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "orf_pxc";
            cTbc.Name = "PxC";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "orf_qta";
            cTbc.Name = "Q.tà";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.BackColor = Color.GreenYellow;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "orf_cos";
            cTbc.Name = "Costo";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "orf_prv";
            cTbc.Name = "Prezzo vendita";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "OrfImp";
            cTbc.Name = "Importo";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            //cTbc.DataPropertyName = "orf_msg";
            cTbc.DataPropertyName = "tmp_msg";
            cTbc.Name = "Messaggi";
            cTbc.Width = 150;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.MaxInputLength = 50;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            //cTbc = new DataGridViewTextBoxColumn();
            //cTbc.DataPropertyName = "orf_arf";
            //cTbc.Name = "Art. fornitore";
            //cTbc.Width = 60;
            //cTbc.ValueType = typeof(string);
            //cTbc.ReadOnly = true;
            ////cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            //dgv1.Columns.Add(cTbc);

            //cTbc = new DataGridViewTextBoxColumn();
            //cTbc.DataPropertyName = "orf_ean";
            //cTbc.Name = "Barcode";
            //cTbc.Width = 70;
            //cTbc.ValueType = typeof(string);
            //cTbc.ReadOnly = true;
            ////cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            //dgv1.Columns.Add(cTbc);

            //cTbc = new DataGridViewTextBoxColumn();
            //cTbc.DataPropertyName = "OrfMdy";
            //cTbc.Name = "Modificato";
            //cTbc.Width = 0;
            //cTbc.ValueType = typeof(string);
            //cTbc.ReadOnly = false;
            ////cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            //dgv1.Columns.Add(cTbc);
        }

        private void FillTabs()
        {            
            DataTable tFor = _clsFun.FillTabSql("AnaFornitori", "SELECT for_cod, for_des FROM AnaFornitori", false, _strConSql);

            cmbOrdFor.DataSource = tFor;
            cmbOrdFor.DisplayMember = "for_des";
            cmbOrdFor.ValueMember = "for_cod";
            cmbOrdFor.SelectedValue = "";
        }

        private void FillDati()
        {
            dgv1.DataSource = _tabOrd;
        }

        private void dgv1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 2 || e.ColumnIndex == 3)
            {

                string sArt = dgv1.Rows[e.RowIndex].Cells["Articolo"].Value.ToString();

                if (e.ColumnIndex == 2 || sArt.Trim() == "")
                {
                    frmSeekArt f = new frmSeekArt();
                    f.ShowDialog();
                    if (f._tabArt != null && f._tabArt.Rows.Count > 0)
                    {
                        DataTable t = f._tabArt;
                        if (t.Rows.Count > 0)
                        {
                            CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                            if (cm.Position >= 0)
                            {
                                DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                                DataRow x = r.Row;

                                //sArt = (string)t.Rows[0]["tmp_art"];

                                //t = _clsQry.ArtSeek(sArt, "SEEK");

                                //t = _clsQry.ArtSeek(sArt, "FOR" + cmbOftFor.SelectedValue.ToString());

                                //FillOrfRow(t, x, true);
                            }
                        }
                    }
                }
                else
                {
                    frmAnaArticolo f = new frmAnaArticolo();
                    f._strArtCod = sArt;
                    f.ShowDialog();
                }
            }
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            if (cmbOrdFor.SelectedValue == null || cmbOrdFor.SelectedValue.ToString() == "")
                MessageBox.Show("Fornitore non delezionato!", "Controllo fornitore", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
                CambioFornitore();
        }

        private void CambioFornitore()
        {
            CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {
                DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                DataRow x = r.Row;

                string s = (string)x["tmp_for"];

                string sFor = cmbOrdFor.SelectedValue.ToString();

                if(sFor == "")
                    MessageBox.Show("Fornitore non definito!", "Controllo fornitore",MessageBoxButtons.OK, MessageBoxIcon.Error);
                else if(sFor == s)
                    MessageBox.Show("Fornitore non cambiato!", "Controllo fornitore", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else 
                {
                    string sArt = (string)x["orf_art"];

                    s = "SELECT * FROM GesLisAcquisto WHERE lia_for='" + sFor + "' AND lia_art='" + sArt + "' ORDER BY lia_dti DESC";
                    DataTable t = _clsFun.FillTabSql("GesLisAcquisto", s, false, _strConSql);

                    if(t.Rows.Count == 0)
                        MessageBox.Show("Fornitore non definito per l'articolo selezionato!", "CONTROLLO FORNITORE", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else
                    {
                        s = (string)t.Rows[0]["lia_arf"];
                        x["tmp_for"] = sFor;
                        x["orf_arf"] = s;
                    }
                }
            }
        }

    }
}
