using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace APOffice
{
    public partial class frmGesVarStorico : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        private const string TABGESVAR = "GesVariazioni";

        private string _strConSql = "";

        public string _strVarTip = "";
        public string _strVarNum = "";

        public frmGesVarStorico()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);;
        }

        private void frmGesVarStorico_Load(object sender, EventArgs e)
        {
            this.Text += " ";
            if (_strVarTip == _clsDef.VARETI)
                this.Text += "Etichette";
            else
                this.Text += "variazioni casse/bilance";

            _strConSql = _clsFun.ConSql(""); 
            SetDgv1();
            FillDati();
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            VarNum();

            this.Close();
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
            cTbc.DataPropertyName = "VarCho";
            cTbc.Name = "Selezione";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            //cTbc.ReadOnly = true;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "var_num";
            cTbc.Name = "Numero";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            //cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "var_dtv";
            cTbc.Name = "Data";
            cTbc.Width = 90;
            cTbc.ValueType = typeof(string);
            //cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "VarCnt";
            cTbc.Name = "Articoli";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            //cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
        }

        private void FillDati()
        {
            string s = "";

            s += "SELECT COUNT(*) AS VarCnt, var_num, var_dtv ";
            s += "FROM GesVariazioni ";
            s += "WHERE (var_tip = '" + _strVarTip + "') AND (var_num <> '000') ";
            s += "GROUP BY var_num, var_dtv ";
            s += "ORDER BY var_dtv DESC, var_num DESC";

            DataTable t = _clsFun.FillTabSql(TABGESVAR, s, false, _strConSql);

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "VarCho",
                Caption = "Cho",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });


            dgv1.DataSource = t;
        }

        private void dgv1_DoubleClick(object sender, EventArgs e)
        {
            //CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, this.dgv1.DataMember] as CurrencyManager;
            //if (cm.Position >= 0)
            //{
            //    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
            //    DataRow x = r.Row;
            //    _strVarNum = (string)x["var_num"];
            //    this.Close();
            //}
        }

        private void dgv1_Click(object sender, EventArgs e)
        {
            Console.WriteLine("zxc");


            if (dgv1.DataSource != null)
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    //string sCod = (string)r.Row["tab_cod"];

                    cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                    if (cm.Position >= 0)
                    {
                        r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                        if ((string)r.Row["VarCho"] == "S")
                            r.Row["VarCho"] = "";
                        else
                            r.Row["VarCho"] = "S";
                    }
                }
            }

        }

        //private void dgv1_CellClick(object sender, DataGridViewCellEventArgs e)
        //{
        //    Console.WriteLine("zxc");

        //    //if (dgv1.DataSource != null)
        //    //{
        //    //    CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
        //    //    if (cm.Position >= 0)
        //    //    {
        //    //        DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
        //    //        string sCod = (string)r.Row["tab_cod"];

        //    //        cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
        //    //        if (cm.Position >= 0)
        //    //        {
        //    //            r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
        //    //            r.Row["tab_cod"] = sCod;
        //    //            r.Row["TabMdy"] = "S";
        //    //        }
        //    //    }
        //    //}
        //}

        private void VarNum()
        {
            DataTable t = (DataTable)dgv1.DataSource;

            foreach(DataRow y in t.Rows)
            {
                if((string)y["VarCho"] == "S")
                    _strVarNum += (string)y["var_num"] + ";";
            }
        }

    }
}
