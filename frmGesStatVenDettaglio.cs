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
    public partial class frmGesStatVenDettaglio : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        public DateTime _dayDti = DateTime.Today;
        public DateTime _dayDtf = DateTime.Today;

        private string _strConSql = "";
        private string _strConSqlSta = "";

        public frmGesStatVenDettaglio()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
            _strConSql = _clsFun.ConSql("");
            _strConSqlSta = _clsFun.ConSql("3");
        }

        private void frmGesStatVenDettaglio_Load(object sender, EventArgs e)
        {
            dtpDti.Value = _dayDti;
            dtpDtf.Value = _dayDtf;
            //SetDgv1();
            FillDati();
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }
        private void frmGesStatVenDettaglio_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
        }
        private void Esci()
        {
            this.Close();
        }

        private void btnEstrai_Click(object sender, EventArgs e)
        {
            FillDati();
        }

        private void FillDati()
        {
            string s = "";
            DataRow x;
            DataRow[] j;
            DataRow[] j2;

            //DataTable t = new clsGenTabTmp().TabTmpVenReparto("TabRep");

            //s = "SELECT * FROM TabReparti";
            //DataTable tRep = _clsFun.FillTabSql("TabRep", s, false, _strConSql);

            //s = "SELECT AnaArticoli.art_cod, AnaArticoli.art_des, AnaArticoli.art_rep, TabReparti.tab_des AS RepDes ";
            //s += "FROM AnaArticoli ";
            //s += "LEFT OUTER JOIN TabReparti ON AnaArticoli.art_rep = TabReparti.tab_cod";
            //DataTable tArt = _clsFun.FillTabSql("Art", s, false, _strConSql);
            //DataColumn[] keys = new DataColumn[1];
            //keys[0] = tArt.Columns["art_cod"];
            //tArt.PrimaryKey = keys;

            s = "SELECT * FROM GesNegVen ";
            s += "WHERE (ven_day >= " + _clsFun.DaySql(dtpDti.Value) + ") AND (ven_day <= " + _clsFun.DaySql(dtpDtf.Value) + ") ";
            //s += "GROUP BY ven_art";
            DataTable tVen = _clsFun.FillTabSql("Sta", s, false, _strConSqlSta);

            //tVen.Columns.Add(new DataColumn()
            //{
            //    DataType = Type.GetType("System.String"),
            //    ColumnName = "tmp_ard",
            //    Caption = "Descrizione",
            //    MaxLength = 50,
            //    ReadOnly = false,
            //    DefaultValue = (String)""
            //});

            //tVen.Columns.Add(new DataColumn()
            //{
            //    DataType = Type.GetType("System.String"),
            //    ColumnName = "tmp_red",
            //    Caption = "Descrizione",
            //    MaxLength = 50,
            //    ReadOnly = false,
            //    DefaultValue = (String)""
            //});


            //foreach (DataRow y in tVen.Rows)
            //{
            //    j2 = tArt.Select("art_cod='" + y["ven_art"] + "'");
            //    if (j2.Length > 0)
            //    {
            //        j = t.Select("vre_rep='" + j2[0]["art_rep"] + "'");
            //        if (j.Length == 0)
            //        {
            //            x = t.NewRow();
            //            x["vre_rep"] = j2[0]["art_rep"];
            //            x["vre_red"] = j2[0]["RepDes"];
            //            x["vre_ven"] = 0;
            //            t.Rows.Add(x);

            //            j = t.Select("vre_rep='" + j2[0]["art_rep"] + "'");

            //        }

            //        j[0]["vre_ven"] = (decimal)j[0]["vre_ven"] + (decimal)y["ven_ven"];
            //    }

            //}

            dgv1.DataSource = tVen;

        }

    }
}
