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
    public partial class frmGesStaTipoPagamenti : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        private string _strConSql = "";
        private string _strConSqlSta = "";

        public DateTime _dayDay = new DateTime(2050, 1, 1, 0, 0, 0);

        public frmGesStaTipoPagamenti()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
        }

        private void frmGesStaTipoPagamenti_Load(object sender, EventArgs e)
        {
            this.Text += " al " + _dayDay.ToString("dd/MM/yyyy");

            _strConSql = _clsFun.ConSql("");
            _strConSqlSta = _clsFun.ConSql("3");

            SetDgv1();
            FillDati();
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }
        private void frmGesStaTipoPagamenti_KeyDown(object sender, KeyEventArgs e)
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
            dgv1.ReadOnly = true;
            dgv1.AllowUserToDeleteRows = false;
            //dgv1.DisplayedRowCount() = true;

            DataGridViewTextBoxColumn cTbc;
            //DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vep_pos";
            cTbc.Name = "Cassa";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vep_cod";
            cTbc.Name = "Codice";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "PagDes";
            cTbc.Name = "Pagamento";
            cTbc.Width = 250;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "VepImp";
            cTbc.Name = "Valore";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
        }

        public void FillDati()
        {
            string s = "";
            DataRow x;
            DataRow[] j;

            s = "SELECT * FROM TabCauCassa";
            DataTable tCau = _clsFun.FillTabSql("Cau", s, false, _strConSql);

            s = "SELECT ";
            s += "vep_pos, vep_cod, SUM(vep_imp) AS VepImp ";
            s += "FROM GesNegVep ";
            s += "WHERE ";
            s += "vep_day = " + _clsFun.DaySql(_dayDay) + " ";
            s += "GROUP BY vep_pos, vep_cod ";
            s += "ORDER BY vep_cod";
            DataTable tVep = _clsFun.FillTabSql("Vep", s, false, _strConSqlSta);

            s = "SELECT * FROM GesNegMca ";
            s += "WHERE mca_day = " + _clsFun.DaySql(_dayDay) + " ";
            DataTable tMca = _clsFun.FillTabSql("Mca", s, false, _strConSqlSta);

            tVep.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "PagCod",
                Caption = "Codice",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (string)""
            });

            tVep.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "PagDes",
                Caption = "Descrizione",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (string)""
            });


            DataTable t = tVep.Clone();

            foreach (DataRow y in tVep.Rows)
            {
                j = tCau.Select("tab_cod='" + (string)y["vep_cod"] + "'");
                if (j.Length > 0)
                {

                    y["PagCod"] = (string)j[0]["tab_ppo"];
                    y["PagDes"] = (string)j[0]["tab_des"];

                    int iSgn = 1;
                    s = (string)y["vep_cod"];

                    if((string)j[0]["tab_key"] == "RES")
                    {
                        iSgn = -1;
                        s = "001";
                    }
                       
                    j = t.Select("vep_pos='" + y["vep_pos"] + "' AND vep_cod='" + s + "'");
                    if(j.Length == 0)
                    {
                        x = t.NewRow();
                        x["vep_pos"] = y["vep_pos"];
                        x["vep_cod"] = s;
                        x["VepImp"] = 0;
                        x["PagCod"] = y["PagCod"];
                        x["PagDes"] = y["PagDes"];
                        t.Rows.Add(x);
                    }

                    j = t.Select("vep_pos='" + y["vep_pos"] + "' AND vep_cod='" + s + "'");
                    j[0]["VepImp"] = (decimal)j[0]["VepImp"] + ((decimal)y["VepImp"] * iSgn);

                }
            }

            tVep = t.Copy();

            foreach (DataRow y in tMca.Rows)
            {
                x = tVep.NewRow();
                x["PagDes"] = y["mca_cau"];
                if ((string)x["PagDes"] == "VER")
                    x["PagDes"] = "Versamento";
                else if ((string)x["PagDes"] == "PRE")
                    x["PagDes"] = "Prelevamento";

                if ((string)y["mca_ora"] != "")
                    x["PagDes"] += "(" + (string)y["mca_ora"] + " - " + (string)y["mca_usr"] + ")";

                x["VepImp"] = y["mca_imp"];
                tVep.Rows.Add(x);
            }

            dgv1.DataSource = tVep;
        }

    }
}
