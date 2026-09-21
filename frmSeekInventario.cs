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
    public partial class frmSeekInventario : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        private const string TABGESINT = "GesInvTestate";
        private const string TABGESINV = "GesInventario";

        private string _strConSql = "";

        public frmSeekInventario()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
        }

        private void frmSeekInventario_Load(object sender, EventArgs e)
        {
            ApplyModernUi();
            _strConSql = _clsFun.ConSql(""); 
            SetDgv1();
            FillDati();

            clsUiIcons.RestoreFormBounds(this);
            clsUiIcons.RestoreGridColumnWidths(dgv1, "frmSeekInventario_dgv1");
        }

        private void frmSeekInventario_FormClosing(object sender, FormClosingEventArgs e)
        {
            clsUiIcons.SaveFormBounds(this);
            clsUiIcons.SaveGridColumnWidths(dgv1, "frmSeekInventario_dgv1");
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
                    if (utilityToolStripMenuItem != null)
                        utilityToolStripMenuItem.Image = clsUiIcons.GetIcon("tools_gear", 16);
                    if (comtrolloNegozioToolStripMenuItem != null)
                        comtrolloNegozioToolStripMenuItem.Image = clsUiIcons.GetIcon("store", 16);
                }

                this.BackColor = Color.FromArgb(243, 244, 246);

                if (btnNew != null)
                {
                    clsUiIcons.StyleStatButton(btnNew, "Nuovo", "plus", 16,
                        Color.FromArgb(236, 253, 245), Color.FromArgb(167, 243, 208),
                        Color.FromArgb(52, 211, 153), Color.FromArgb(6, 78, 59), Color.FromArgb(5, 150, 105));
                }

                if (btnCan != null)
                {
                    clsUiIcons.StyleStatButton(btnCan, "Elimina", "trash", 16,
                        Color.FromArgb(254, 242, 242), Color.FromArgb(254, 202, 202),
                        Color.FromArgb(248, 113, 113), Color.FromArgb(127, 29, 29), Color.FromArgb(220, 38, 38));
                }

                if (dgv1 != null)
                {
                    clsUiIcons.StyleDataGridView(dgv1);
                }
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("frmSeekInventario.ApplyModernUi", ex.Message);
            }
        }

        private void frmSeekInventario_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Escape)
                this.Close();
        }
        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }
        private void Esci()
        {
            this.Close();
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            frmGesInventario f = new frmGesInventario();
            f._strInvNum = _clsDef.CODNEW;
            f.ShowDialog();
            FillDati();
        }
        private void FillDati()
        {
            string s = "";
            s = "SELECT * FROM GesInvTestate LEFT OUTER JOIN TabReparti ON GesInvTestate.int_rep = TabReparti.tab_cod ";
            s += "ORDER BY int_num,int_day DESC";
            DataTable t = _clsFun.FillTabSql(TABGESINT, s, false, _strConSql);

            dgv1.DataSource = t;
        }

        private void SetDgv1()
        {
            dgv1.AutoGenerateColumns = false;
            //dgv1.VirtualMode = true;
            //dgv1.Dock = DockStyle.Fill;
            dgv1.AllowUserToAddRows = false;
            dgv1.ReadOnly = true;
            dgv1.AllowUserToDeleteRows = false;
            //dgv1.DisplayedRowCount() = true;

            DataGridViewTextBoxColumn cTbc;
            //DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "int_day";
            cTbc.Name = "Data";
            cTbc.Width = 80;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "int_neg";
            cTbc.Name = "Negozio";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "int_num";
            cTbc.Name = "Numero";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "int_tip";
            cTbc.Name = "Tipo";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "int_des";
            cTbc.Name = "Note";
            cTbc.Width = 300;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 50;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tab_des";
            cTbc.Name = "Reparto";
            cTbc.Width = 300;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 50;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);
        }

        private void dgv1_DoubleClick(object sender, EventArgs e)
        {
            Scelto();
        }

        private void Scelto()
        {
            CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, this.dgv1.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {
                DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                DataRow x = r.Row;

                frmGesInventario f = new frmGesInventario();
                f._strInvNum = (string)x["int_num"];
                f._strInvNeg = (string)x["int_neg"];
                f.ShowDialog();

                if (f._strInvDes != "")
                    x["int_des"] = f._strInvDes.Trim();
            }
        }

        private void btnCan_Click(object sender, EventArgs e)
        {
            CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {
                string s = "";
                DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                DataRow x = r.Row;

                if(_clsFun.Msg("SINO",false,"CONFERMI LA CANCELLAZIONE DELL'INVENTARIO DEL " + ((DateTime)x["int_day"]).ToString("dd/MM/yyy")+"?") == "SI")
                {
                    s = "DELETE FROM " + TABGESINT + " WHERE int_num='" + (string)x["int_num"] + "'";
                    _clsFun.SqlWrite(s, _strConSql);
                    s = "DELETE FROM " + TABGESINV + " WHERE inv_num='" + (string)x["int_num"] + "'";
                    _clsFun.SqlWrite(s, _strConSql);

                    FillDati();
                }
            }
        }

        private void comtrolloNegozioToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new frmUtyInvNegozio().ShowDialog();
        }

        private void utilityToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void dgv1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

    }
}
