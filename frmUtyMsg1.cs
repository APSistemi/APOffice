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
    public partial class frmUtyMsg1 : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        public string _strTip = "001";
        public string _strStr = "";
        public string _strMsg = "";

        public frmUtyMsg1()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0); ;
        }

        private void frmUtyMsg1_Load(object sender, EventArgs e)
        {
            SetDgv1();
            FillDatiFromString();
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void frmUtyMsg1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
        }
        
        private void Esci()
        {
            this.Close();
        }

        private void SetDgv1()
        {
            dgv1.AutoGenerateColumns = false;
            //dgv1.VirtualMode = true;
            //dgv1.Dock = DockStyle.Fill;
            dgv1.AllowUserToAddRows = false;
            dgv1.ReadOnly = false;
            dgv1.AllowUserToDeleteRows = false;
            //dgv1.DisplayedRowCount() = true;

            DataGridViewTextBoxColumn cTbc;
            //DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_msg";
            cTbc.Name = "Messaggio";
            cTbc.Width = 250;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);
        }

        private void FillDatiFromString()
        {
            DataTable t = new DataTable("tab");
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_msg",
                Caption = "Descrizione",
                MaxLength = 100,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            DataRow x;
            string[] a = _strStr.Split(';');

            for(int i = 0; i < a.Length; i++)
            {
                x = t.NewRow();
                x["tmp_msg"] = a[i];
                t.Rows.Add(x);
            }

            dgv1.DataSource = t;
        }

        private void dgv1_DoubleClick(object sender, EventArgs e)
        {
            Scelto();
        }

        private void dgv1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
                Scelto();
        }

        private void Scelto()
        {
            if (dgv1.DataSource != null)
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, this.dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    _strMsg = (string)x["tmp_msg"];

                    Esci();
                }
            }
        }


    }


}
