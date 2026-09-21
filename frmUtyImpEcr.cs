using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace APOffice
{
    public partial class frmUtyImpEcr : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        private const string TABLV1 = "TabEcrLv1";
        private const string TABLV2 = "TabEcrLv2";
        private const string TABLV3 = "TabEcrLv3";
        private const string TABREP = "TabReparti";

        private string _strConSql = "";

        public frmUtyImpEcr()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);;
        }

        private void frmUtyImpEcr_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");
            SetDgv2();
            SetDgv3();
            SetDgv4();
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void SetDgv2()
        {
            dgv2.AutoGenerateColumns = false;
            //dgv1.VirtualMode = true;
            //dgv1.Dock = DockStyle.Fill;
            dgv2.AllowUserToAddRows = false;
            dgv2.ReadOnly = false;
            dgv2.AllowUserToDeleteRows = false;
            //dgv1.DisplayedRowCount() = true;

            DataGridViewTextBoxColumn cTbc;
            //DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tab_cod";
            cTbc.Name = "Codice";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tab_des";
            cTbc.Name = "Descrizione";
            cTbc.Width = 190;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv2.Columns.Add(cTbc);
        }

        private void SetDgv3()
        {
            dgv3.AutoGenerateColumns = false;
            //dgv1.VirtualMode = true;
            //dgv1.Dock = DockStyle.Fill;
            dgv3.AllowUserToAddRows = false;
            dgv3.ReadOnly = false;
            dgv3.AllowUserToDeleteRows = false;
            //dgv1.DisplayedRowCount() = true;

            DataGridViewTextBoxColumn cTbc;
            //DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tab_lv1";
            cTbc.Name = "Lv1";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv3.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tab_cod";
            cTbc.Name = "Codice";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv3.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tab_des";
            cTbc.Name = "Descrizione";
            cTbc.Width = 190;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv3.Columns.Add(cTbc);
        }

        private void SetDgv4()
        {
            dgv4.AutoGenerateColumns = false;
            //dgv1.VirtualMode = true;
            //dgv1.Dock = DockStyle.Fill;
            dgv4.AllowUserToAddRows = false;
            dgv4.ReadOnly = false;
            dgv4.AllowUserToDeleteRows = false;
            //dgv1.DisplayedRowCount() = true;

            DataGridViewTextBoxColumn cTbc;
            //DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tab_lv2";
            cTbc.Name = "Lv2";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv4.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tab_lv1";
            cTbc.Name = "Lv1";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv4.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tab_cod";
            cTbc.Name = "Cod";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv4.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tab_des";
            cTbc.Name = "Descrizione";
            cTbc.Width = 190;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv4.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tab_rep";
            cTbc.Name = "Reparto";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv4.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tab_key";
            cTbc.Name = "Key";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv4.Columns.Add(cTbc);
        }

        private void btnMdb_Click(object sender, EventArgs e)
        {
            OpenFileDialog o = new OpenFileDialog();
            //openFile1.InitialDirectory = "C:\\Programmi\\XFOODNET\\DB\\";
            o.InitialDirectory = "C:\\ApProject\\Temp\\";
            o.Filter = "XLS Files|*.xls|All Files(*.*)|*.*";

            if (o.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                lblXls.Text = o.FileName;

            if (Path.GetExtension(lblXls.Text).ToLower() == ".xls")
                FillXls(lblXls.Text);
        }

        private void FillXls( string strXls)
        {
            OleDbConnection cn = new System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + strXls + "';Extended Properties=Excel 8.0;");
            cn.Open();
            OleDbCommand cm = new OleDbCommand();
            cm.Connection = cn;

            DataTable sch = cn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            string sTab = (string)sch.Rows[0][2]; 

            //string s = "SELECT * FROM [Sheet1$]";
            string s = "SELECT * FROM [" + sTab + "]";

            cm.CommandText = s;
            OleDbDataReader dr = cm.ExecuteReader();
            DataTable t = new DataTable();
            t.TableName = "TabXls";
            t.Load(dr);
            dgv1.DataSource = t;
        }

        private void btnFill_Click(object sender, EventArgs e)
        {
            FillEcr();
        }

        private void FillEcr()
        {
            string s = "";
            DataRow[] j;
            DataRow x;

            //((DataTable)dgv1.DataSource).Clear();
            //((DataTable)dgv2.DataSource).Clear();
            if(dgv2.DataSource != null)
                ((DataTable)dgv2.DataSource).Clear();


            int iLv1 = Convert.ToInt16(txtPosCodLv1.Text);
            int iLl1 = Convert.ToInt16(txtLenCodLv1.Text);
            int iLd1 = Convert.ToInt16(txtPosDesLv1.Text);

            int iLv2 = Convert.ToInt16(txtPosCodLv2.Text);
            int iLl2 = Convert.ToInt16(txtLenCodLv2.Text);
            int iLd2 = Convert.ToInt16(txtPosDesLv2.Text);

            int iLv3 = Convert.ToInt16(txtPosCodLv3.Text);
            int iLl3 = Convert.ToInt16(txtLenCodLv3.Text);
            int iLd3 = Convert.ToInt16(txtPosDesLv3.Text);

            int iRep = Convert.ToInt16(txtPosCodRep.Text);
            int iRel = Convert.ToInt16(txtLenCodRep.Text);

            s = "SELECT * FROM " + TABREP;
            DataTable tRep = _clsFun.FillTabSql(TABREP, s, false, _strConSql);

            s = "SELECT * FROM " + TABLV1;
            DataTable tLv1 = _clsFun.FillTabSql(TABLV1, s, false, _strConSql);

            s = "SELECT * FROM " + TABLV2;
            DataTable tLv2 = _clsFun.FillTabSql(TABLV2, s, false, _strConSql);

            s = "SELECT * FROM " + TABLV3;
            DataTable tLv3 = _clsFun.FillTabSql(TABLV3, s, false, _strConSql);

            foreach(DataRow y in ((DataTable)dgv1.DataSource).Rows)
            {
                try
                {
                    string sLv1 = Convert.ToString(y[iLv1 - 1]).PadLeft(iLl1, Convert.ToChar("0"));
                    string sLd1 = Convert.ToString(y[iLd1 - 1]);
                    if (sLd1.Length > 30)
                        sLd1 = sLd1.Substring(0,30);

                    string sLv2 = Convert.ToString(y[iLv2 - 1]).PadLeft(iLl2, Convert.ToChar("0"));
                    string sLd2 = Convert.ToString(y[iLd2 - 1]);
                    if (sLd2.Length > 30)
                        sLd2 = sLd2.Substring(0, 30);


                    string sLv3 = Convert.ToString(y[iLv3 - 1]).PadLeft(iLl3, Convert.ToChar("0"));
                    string sLd3 = Convert.ToString(y[iLd3 - 1]);
                    if (sLd3.Length > 30)
                        sLd3 = sLd3.Substring(0, 30);

                    if (sLv1 != "001" && sLv1 != "002" && sLv1 != "000")
                        Console.WriteLine("zzzzzzz");

                    string sRep = "";
                    if(iRel > 0)
                        sRep = Convert.ToString(y[iRep - 1]).PadLeft(iRel, Convert.ToChar("0"));
                    if (sRep.Length > 3)
                        sRep = sRep.Substring(1, 3);

                    string sKey = Convert.ToString(y[1]) + Convert.ToString(y[4]);

                    if (sLv1 != "000" && sLv2 != "000" && sLv3 != "000")
                    {
                        j = tLv1.Select("tab_cod='" + sLv1 + "'");
                        if (j.Length == 0)
                        {
                            x = tLv1.NewRow();
                            x["tab_cod"] = sLv1;
                            x["tab_des"] = sLd1;
                            tLv1.Rows.Add(x);
                        }

                        j = tLv2.Select("tab_cod='" + sLv2 + "' AND tab_lv1='" + sLv1 + "'");
                        if (j.Length == 0)
                        {
                            x = tLv2.NewRow();
                            x["tab_cod"] = sLv2;
                            x["tab_lv1"] = sLv1;
                            x["tab_des"] = sLd2;
                            tLv2.Rows.Add(x);
                        }

                        j = tLv3.Select("tab_cod='" + sLv3 + "' AND tab_lv2='" + sLv2 + "' AND tab_lv1='" + sLv1 + "'");
                        if (j.Length == 0)
                        {
                            x = tLv3.NewRow();
                            x["tab_cod"] = sLv3;
                            x["tab_lv2"] = sLv2;
                            x["tab_lv1"] = sLv1;
                            x["tab_des"] = sLd3;

                            x["tab_rep"] = "";
                            j = tRep.Select("tab_cod='" + sRep + "'");
                            if (j.Length > 0)
                                x["tab_rep"] = sRep;

                            x["tab_key"] = sKey;

                            tLv3.Rows.Add(x);
                        }
                    }
                }
                catch (Exception ex)
                {

                    if (MessageBox.Show(ex.Message, "USCIRE?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        break;
                    //this._clsMail.InvMail("edp@spacrovigo.org", "EDP", "paolo.secchettin@spacrovigo.org", "", "Errori su SpacServer", "clsGenFun.SqlWrite", "", "");
                }
            }

            dgv2.DataSource = tLv1;
            dgv3.DataSource = tLv2;
            dgv4.DataSource = tLv3;

        }

        private void button1_Click(object sender, EventArgs e)
        {
            //string s = "";
            //DataRow[] j;
            //DataRow x;

            //int iCod = Convert.ToInt16(txtPosCodLv1.Text);
            //int iDes = Convert.ToInt16(txtDes.Text);
            //int iLv1 = Convert.ToInt16(txtLenCodLv1.Text);
            //int iLv2 = Convert.ToInt16(txtLv2.Text);
            //int iLv3 = Convert.ToInt16(txtLv3.Text);
            //int iRep = Convert.ToInt16(txtRep.Text);

            //iLv1 = 0;
            //iLv2 = 2;
            //iLv3 = 4;

            //s = "SELECT * FROM " + TABREP;
            //DataTable tRep = _clsFun.FillTabSql(TABREP, s, false, _strConSql);

            //s = "SELECT * FROM " + TABLV1;
            //DataTable tLv1 = _clsFun.FillTabSql(TABLV1, s, false, _strConSql);

            //s = "SELECT * FROM " + TABLV2;
            //DataTable tLv2 = _clsFun.FillTabSql(TABLV2, s, false, _strConSql);

            //s = "SELECT * FROM " + TABLV3;
            //DataTable tLv3 = _clsFun.FillTabSql(TABLV3, s, false, _strConSql);

            //foreach (DataRow y in ((DataTable)dgv1.DataSource).Rows)
            //{
            //    string sCod = Convert.ToString(y[iCod - 1]);
            //    string sDes = (string)y[iDes - 1];
            //    //string sRep = Convert.ToString(y[iRep - 1]).PadLeft(3, Convert.ToChar("0"));

            //    string sLv1 = Convert.ToInt16(y[iLv1]).ToString("000");
            //    string sLv2 = Convert.ToInt16(y[iLv2]).ToString("000");
            //    string sLv3 = Convert.ToInt16(y[iLv3]).ToString("000");

            //    string sLd1 = ((string)y[iLv1 + 1]).Trim();
            //    string sLd2 = ((string)y[iLv2 + 1]).Trim();
            //    string sLd3 = ((string)y[iLv3 + 1]).Trim();

            //    //if (sLv2 == "000")
            //    //{
            //        j = tLv1.Select("tab_cod='" + sLv1 + "'");
            //        if (j.Length == 0)
            //        {
            //            x = tLv1.NewRow();
            //            x["tab_cod"] = sLv1;
            //            x["tab_des"] = sLd1;
            //            tLv1.Rows.Add(x);
            //        }
            //    //}
            //    //else if (sLv3 == "000")
            //    //{
            //        j = tLv2.Select("tab_cod='" + sLv2 + "' AND tab_lv1='" + sLv1 + "'");
            //        if (j.Length == 0)
            //        {
            //            x = tLv2.NewRow();
            //            x["tab_cod"] = sLv2;
            //            x["tab_lv1"] = sLv1;
            //            x["tab_des"] = sLd2;
            //            tLv2.Rows.Add(x);
            //        }
            //    //}
            //    //else
            //    //{
            //        j = tLv3.Select("tab_cod='" + sLv3 + "' AND tab_lv2='" + sLv2 + "' AND tab_lv1='" + sLv1 + "'");
            //        if (j.Length == 0)
            //        {
            //            x = tLv3.NewRow();
            //            x["tab_cod"] = sLv3;
            //            x["tab_lv2"] = sLv2;
            //            x["tab_lv1"] = sLv1;
            //            x["tab_des"] = sLd3;

            //            //x["tab_rep"] = "";
            //            //j = tRep.Select("tab_cod='" + sRep + "'");
            //            //if (j.Length > 0)
            //            //    x["tab_rep"] = sRep;

            //            tLv3.Rows.Add(x);
            //        }
            //    //}
            //}

            //dgv2.DataSource = tLv1;
            //dgv3.DataSource = tLv2;
            //dgv4.DataSource = tLv3;

        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            Salva();
        }

        private void Salva()
        {
            string s = "";
            DataRow[] j;

            DataTable t = new DataTable("TabTmp");
            DataTable tLiv = new DataTable("TabLiv");

            //string sTab = "";

            ArrayList aWhe = new ArrayList();
            ArrayList aExl = new ArrayList();

            for(int i = 1; i <= 3; i++)
            {
                if (i == 1)
                {
                    t = (DataTable)dgv2.DataSource;

                    s = "SELECT * FROM " + TABLV1;
                    tLiv = _clsFun.FillTabSql(TABLV1, s, false, _strConSql);
                    aWhe = new ArrayList();
                    aWhe.Add("tab_cod");
                }
                else if (i == 2)
                {
                    t = (DataTable)dgv3.DataSource;

                    s = "SELECT * FROM " + TABLV2;
                    tLiv = _clsFun.FillTabSql(TABLV2, s, false, _strConSql);
                    aWhe = new ArrayList();
                    aWhe.Add("tab_cod");
                    aWhe.Add("tab_lv1");
                }
                else if (i == 3)
                {
                    t = (DataTable)dgv4.DataSource;

                    s = "SELECT * FROM " + TABLV3;
                    tLiv = _clsFun.FillTabSql(TABLV3, s, false, _strConSql);
                    aWhe = new ArrayList();
                    aWhe.Add("tab_cod");
                    aWhe.Add("tab_lv1");
                    aWhe.Add("tab_lv2");
                }

                foreach(DataRow y in t.Rows)
                {
                    string sSql = "tab_cod='" + (string)y["tab_cod"] + "'";
                    if (i == 2 || i == 3)
                        sSql += " AND tab_lv1='" + (string)y["tab_lv1"] + "'";
                    if (i == 3)
                        sSql += " AND tab_lv2='" + (string)y["tab_lv2"] + "'";
                    j = tLiv.Select(sSql);
                    if (j.Length == 0)
                        s = _clsFun.SqlInsertRow(tLiv.TableName, tLiv, y);
                    else
                        s = _clsFun.SqlUpdRow(tLiv.TableName, tLiv, j[0], y, aWhe, aExl);
                    if (s != "")
                        _clsFun.SqlWrite(s, _strConSql);
                }
            }

            MessageBox.Show("Fine aggiornamento");
        }
    }
}
