using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//using Microsoft.Reporting.WinForms;

namespace APOffice
{
    public partial class frmGesPrn1 : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        public DataTable _tabPrn = new DataTable("tabTmp");
        public DataSet _dasPrn = new DataSet();

        public frmGesPrn1()
        {
            InitializeComponent();
        }

        private void frmGesPrn1_Load(object sender, EventArgs e)
        {
            if (_tabPrn.TableName == "TabRep")
                prnRep(_tabPrn);
            if (_tabPrn.TableName == "TabRep")
                prnFatFor(_tabPrn);

        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        public void prnRep(DataTable tabTab)
        {
            //Stubbed due to missing ReportViewer reference
            /*
            DataSet ds = new DataSet("DataSet1");
            ds.Tables.Add(tabTab);

            ReportDataSource rds = new ReportDataSource();
            rds.Name = "DataSet1";
            rds.Value = ds.Tables[tabTab.TableName];

            reportViewer1.LocalReport.ReportPath = _clsDef.PATHREPORT + "prnRep1.rdlc";
            reportViewer1.LocalReport.DataSources.Clear();
            reportViewer1.LocalReport.DataSources.Add(rds);
            reportViewer1.RefreshReport();
            */
        }

        public void prnFatFor(DataTable tabTab)
        {
            //Stubbed due to missing ReportViewer reference
            /*
            DataSet ds = new DataSet("DataSet1");
            ds.Tables.Add(tabTab.Copy());

            ReportDataSource rds = new ReportDataSource();
            rds.Name = "DataSet1";
            rds.Value = ds.Tables[tabTab.TableName];

            reportViewer1.LocalReport.ReportPath = _clsDef.PATHREPORT + "prnRep1.rdlc";
            reportViewer1.LocalReport.DataSources.Clear();
            reportViewer1.LocalReport.DataSources.Add(rds);
            reportViewer1.RefreshReport();
            */
        }

    }
}
