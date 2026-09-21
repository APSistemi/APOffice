using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Globalization;

namespace APOffice
{
    public partial class frmUtyStaStoCalendario : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        private string _strConSql = "";
        private string _strConSqlSto = "";

        CultureInfo ci = System.Globalization.CultureInfo.CreateSpecificCulture("it");
        DateTimeFormatInfo _ciFormat = new DateTimeFormatInfo();
        Calendar _ciCalendar = CultureInfo.InvariantCulture.Calendar;

        public frmUtyStaStoCalendario()
        {
            InitializeComponent();
            _strConSqlSto = _clsFun.ConSql("5");
        }

        private void frmUtyStaStoCalendario_Load(object sender, EventArgs e)
        {
            _ciFormat = ci.DateTimeFormat;
            _ciCalendar = ci.Calendar;

            cmbYea.Items.Add("");
            for (int i = DateTime.Today.Year - 10; i < DateTime.Today.Year + 1; i++)
                cmbYea.Items.Add(i.ToString());
            cmbYea.SelectedIndex = 0;

            dgv1.DataSource = FillDati();
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }

        private void frmUtyStaStoCalendario_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
        }

        private void Esci()
        {
            this.Close();
        }

        private DataTable FillDati()
        {
            SetDgv1();

            _clsFun.ErrorLog("", _strConSqlSto);

            string s = "SELECT * FROM TabCalendario ORDER BY tab_yea, tab_tri, tab_per";
            DataTable t = _clsFun.FillTabSql("TabCalendario", s, false, _strConSqlSto);
            return t;
        }

        private void btnFill_Click(object sender, EventArgs e)
        {
            if (cmbYea.Text == "")
                MessageBox.Show("ANNO non selezionato!", "CONTROLLO ESTRAZIONE", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            else
            {
                if (MessageBox.Show("Estrazione anno " + cmbYea.Text + ", confermi?", this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
                {
                    FillCalendario(cmbYea.Text);
                }
            }
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
            dgv1.SelectionMode = DataGridViewSelectionMode.CellSelect;

            DataGridViewTextBoxColumn cTbc;
            //DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tab_sta";
            cTbc.Name = "Stato";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tab_yea";
            cTbc.Name = "Anno";
            cTbc.Width = 90;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.Selected = true;
            //cTbc.ToolTipText = "Mensile/Settimanale";
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tab_tri";
            cTbc.Name = "Tipo";
            cTbc.Width = 90;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.Selected = true;
            cTbc.ToolTipText = "Mensile/Settimanale";
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tab_dti";
            cTbc.Name = "D.inizio";
            cTbc.Width = 90;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.Selected = true;
            //cTbc.ToolTipText = "Mensile/Settimanale";
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tab_dtf";
            cTbc.Name = "D.fine";
            cTbc.Width = 90;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.Selected = true;
            //cTbc.ToolTipText = "Mensile/Settimanale";
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tab_per";
            cTbc.Name = "Numero periodo";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.Selected = true;
            //cTbc.ToolTipText = "Mensile/Settimanale";
            dgv1.Columns.Add(cTbc);
        }

        private void FillCalendario(string strYea)
        {
            DataTable t = (DataTable)dgv1.DataSource;
            DataRow x;
            int i = 0;

            DateTime dDay = new DateTime(Convert.ToInt16(strYea), 1, 1);

            DateTime dIni = dDay;
            DateTime dFin = new DateTime();

            int iSet = _ciCalendar.GetWeekOfYear(dDay, _ciFormat.CalendarWeekRule, _ciFormat.FirstDayOfWeek);
            int iSetOld = iSet;

            if (iSet > 50)
            {
                i = 0;
                while (true)
                {
                    i++;
                    dDay = dDay.AddDays(i);
                    iSet = _ciCalendar.GetWeekOfYear(dDay, _ciFormat.CalendarWeekRule, _ciFormat.FirstDayOfWeek);
                    if (iSet == 1)
                    {
                        dIni = dDay;
                        iSetOld = iSet;
                        break;
                    }
                }
            }

            i = 0;
            while (true)
            {
                iSet = _ciCalendar.GetWeekOfYear(dDay, _ciFormat.CalendarWeekRule, _ciFormat.FirstDayOfWeek);

                if (iSet != iSetOld)
                {
                    x = t.NewRow();
                    x["tab_tri"] = "S";
                    x["tab_yea"] = dIni.Year.ToString();    // strYea;
                    x["tab_per"] = iSetOld.ToString("00");
                    x["tab_dti"] = dIni;
                    x["tab_dtf"] = dFin;
                    x["tab_sta"] = "";
                    t.Rows.Add(x);
                    dIni = dDay; //.AddDays(1);

                    iSetOld = iSet;
                }
                dFin = dDay;
                i++;
                dDay = dDay.AddDays(1);

                if (dDay.Year.ToString() != strYea && iSet < 50 && iSet > 1)
                    break;
            }

            for (int iMon = 1; iMon <= 12; iMon++)
            {
                x = t.NewRow();
                x["tab_tri"] = "M";
                x["tab_yea"] = strYea;
                x["tab_per"] = iMon.ToString("00");
                x["tab_dti"] = new DateTime(Convert.ToInt16(strYea), iMon, 1);
                x["tab_dtf"] = _clsFun.FineMese((DateTime)x["tab_dti"]);
                x["tab_sta"] = "";
                t.Rows.Add(x);
            }
        }

        private void dgv1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 0)
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    if ((string)x["tab_sta"] == "S")
                        x["tab_sta"] = "";
                    else
                        x["tab_sta"] = "S";
                }
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            Salva();
        }

        private void Salva()
        {
            string s = "";
            DataRow x;
            DataRow[] j;
            DataTable tCal = FillDati();

            DataTable t = (DataTable)dgv1.DataSource;

            ArrayList aExl = new ArrayList();

            foreach (DataRow y in t.Rows)
            {

                if ((DateTime)y["tab_dti"] == new DateTime(2019, 7, 1))
                    Console.WriteLine("zzzzzzzzz");

                x = tCal.NewRow();
                x["tab_tri"] = y["tab_tri"];
                x["tab_yea"] = y["tab_yea"];
                x["tab_per"] = y["tab_per"];
                x["tab_dti"] = y["tab_dti"];
                x["tab_dtf"] = y["tab_dtf"];
                x["tab_sta"] = y["tab_sta"];

                j = tCal.Select("tab_tri='" + y["tab_tri"] + "' AND tab_yea='" + y["tab_yea"] + "' AND tab_per='" + y["tab_per"] + "'");
                if (j.Length > 0)
                {
                    //s = _clsFun.SqlUpdRow("TabCalendario", tCal, y, x, aWhe, aExl);

                    y["tab_idx"] = j[0]["tab_idx"];
                    s = _clsFun.SqlUpdRowIdx("TabCalendario", tCal, j[0], x, aExl);
                }
                else
                    s = _clsFun.SqlInsertRow("TabCalendario", tCal, x);
                if (s != "")
                    _clsFun.SqlWrite(s, _strConSqlSto);
            }

            MessageBox.Show("Fine aggiornamneto!");
        }

        private void lblSta_Click(object sender, EventArgs e)
        {
            if (lblSta.Text == "")
                lblSta.Text = "S";
            else
                lblSta.Text = "";
        }

        private void btnSta_Click(object sender, EventArgs e)
        {
            AggStato();
        }

        private void AggStato()
        {
            DateTime dDti = dtpDti.Value;
            DateTime dDtf = dtpDtf.Value;
            string sSta = lblSta.Text;

            DataTable t = (DataTable)dgv1.DataSource;

            //TimeSpan diff = dFatDay.Subtract(dd);
            //if (diff.Days != 0 || sPos != sFatPos || sSco != sFatSco)
            //{
            //    b = false;
            //}

            foreach(DataRow y in t.Rows)
            {
                if ((DateTime)y["tab_dti"] >= dDti && (DateTime)y["tab_dtf"] <= dDtf)
                    y["tab_sta"] = sSta;

            }


        }


    }
}
