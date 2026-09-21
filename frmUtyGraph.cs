using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace APOffice
{
    public partial class frmUtyGraph : Form
    {
        private const string TABFON = "TabFonts";

        private clsDefine _clsDef = new clsDefine();
        private clsFuncs _clsFun = new clsFuncs();

        private string _strConSql = "";

        public frmUtyGraph()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);;
        }

        private void frmUtyFonts_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");
            SetDgv1();
            FillDati();
        }

        private void frmUtyGraph_KeyDown(object sender, KeyEventArgs e)
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
            Salva();
            this.Close();
        }


        private void SetDgv1()
        {
            dgv1.AutoGenerateColumns = false;
            dgv1.AllowUserToAddRows = false;
            dgv1.ReadOnly = false;
            dgv1.AllowUserToDeleteRows = false;

            DataGridViewTextBoxColumn cTbc;
            //DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;
            //DataGridViewButtonColumn cBtn;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.ValueType = typeof(bool);
            cTbc.DataPropertyName = "tab_cod";
            cTbc.Name = "Codice";
            cTbc.Width = 30;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tab_des";
            cTbc.Name = "Descrizione";
            cTbc.Width = 100;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tab_fo1";
            cTbc.Name = "Font 1";
            cTbc.Width = 200;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tab_col";
            cTbc.Name = "Color";
            cTbc.Width = 200;
            cTbc.MaxInputLength = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            //c = new DataGridViewTextBoxColumn();
            //c.DataPropertyName = "tab_fo2";
            //c.Name = "Font 2";
            //c.Width = 80;
            //c.ValueType = typeof(string);
            //c.ReadOnly = true;
            //this.dgv1.Columns.Add(c);
            //c = new DataGridViewTextBoxColumn();
            //c.DataPropertyName = "tab_fo3";
            //c.Name = "Font 3";
            //c.Width = 80;
            //c.ValueType = typeof(string);
            //c.ReadOnly = true;
            //this.dgv1.Columns.Add(c);
            //c = new DataGridViewTextBoxColumn();
            //c.DataPropertyName = "tab_fo4";
            //c.Name = "Font 4";
            //c.Width = 80;
            //c.ValueType = typeof(string);
            //c.ReadOnly = true;
            //this.dgv1.Columns.Add(c);
            //c = new DataGridViewTextBoxColumn();
            //c.DataPropertyName = "tab_fo5";
            //c.Name = "Font 5";
            //c.Width = 80;
            //c.ValueType = typeof(string);
            //c.ReadOnly = true;
            //this.dgv1.Columns.Add(c);
        }

        private string ChoFont(Label lbl)
        {
            string s = "";
            fontDialog1.ShowColor = true;
            fontDialog1.Font = new clsGesGraph().Str2Font(new Font("Times New Roman", 7f), lbl.Text);
            fontDialog1.Color = lbl.ForeColor;
            if (fontDialog1.ShowDialog() != DialogResult.Cancel)
            {
                object[] name = new object[] { fontDialog1.Font.FontFamily.Name, ":", fontDialog1.Font.Size, ":", (int)fontDialog1.Font.Style, ":", null };
                Color color = fontDialog1.Color;
                name[6] = color.Name.ToString();
                s = string.Concat(name);
            }
            return s;
        }

        private string ChoColor(Label lbl)
        {
            string s = "";
            //colorDialog1.ShowColor = true;
            //colorDialog1.Font = _clsFun.Str2Font(new Font("Times New Roman", 7f), lbl.Text);
            //colorDialog1.Color = lbl.ForeColor;
            if (colorDialog1.ShowDialog() != DialogResult.Cancel)
            {
                //object[] name = new object[] { colorDialog1.Font.FontFamily.Name, ":", colorDialog1.Font.Size, ":", (int)colorDialog1.Font.Style, ":", null };
                //Color color = colorDialog1.Color;
                //name[6] = color.Name.ToString();
                //s = string.Concat(name);

                Color c = colorDialog1.Color;

                s = System.Drawing.ColorTranslator.ToHtml(c);

                //s = colorDialog1.Color.Name;

            }
            return s;
        }

        private void FillDati(string strCod, bool bolMod)
        {
            string s = "SELECT * FROM TabFonts WHERE tab_cod='" + strCod;
            DataTable t = _clsFun.FillTabSql("TabFonts", s, false, _strConSql);
        }

        private bool Salva()
        {
            dgv1.CommitEdit(DataGridViewDataErrorContexts.Commit);
            bool b = true;
            string s = "SELECT * FROM TabFonts";
            DataTable tOld = _clsFun.FillTabSql("TabFonts", s, false, _strConSql);
            DataTable t = (DataTable)this.dgv1.DataSource;
            foreach (DataRow x in t.Rows)
            {
                DataRow[] j = tOld.Select("tab_cod='" + (string)x["tab_cod"] + "'");
                if ((int)j.Length > 0)
                {
                    ArrayList aWhe = new ArrayList();
                    aWhe.Add("tab_cod");
                    ArrayList aExl = new ArrayList();
                    s = _clsFun.SqlUpdRow("TabFonts", tOld, j[0], x, aWhe, aExl);
                }
                if (s != "")
                {
                    _clsFun.SqlWrite(s, _strConSql);
                }
            }
            return b;
        }

        private void FillDati()
        {
            string s = "SELECT * FROM TabFonts WHERE tab_cod < '100' ORDER BY tab_cod";
            DataTable t = _clsFun.FillTabSql("TabFonts", s, false, _strConSql);
            dgv1.DataSource = t;
        }

        private void dgv1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 2)
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;
                    string s = (string)x["tab_fo1"];
                    Label l = new Label();
                    l.Text = Convert.ToString(x["tab_fo1"]);
                    s = ChoFont(l);
                    if (s != "")
                        x["tab_fo1"] = s;
                }
            }
            if (e.ColumnIndex == 3)
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;
                    string s = (string)x["tab_col"];
                    Label l = new Label();
                    l.BackColor = Color.FromName(s);
                    s = ChoColor(l);
                    if (s != "")
                    {
                        x["tab_col"] = s;
                        if (s.Substring(0, 1) == "#")
                            s = s.Substring(1);
                        lblCol.BackColor = System.Drawing.ColorTranslator.FromHtml("#" + s);
                    }
                }
            }
        }

        private void dgv1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 3)
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;
                    string s = (string)x["tab_col"];
                    if (s.Substring(0, 1) == "#")
                        s = s.Substring(1);
                    lblCol.BackColor = System.Drawing.ColorTranslator.FromHtml("#" + s);

                    //Label l = new Label();
                    //l.BackColor = Color.FromName(s);
                    //s = ChoColor(l);
                    //if (s != "")
                    //    x["tab_col"] = s;
                }
            }

        }

    }
}
