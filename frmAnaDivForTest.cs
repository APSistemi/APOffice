using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace APOffice
{
    public partial class frmAnaDivForTest : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        public DataRow _rowTes;
        public DataTable _tabTab = new DataTable();

        public frmAnaDivForTest()
        {
            InitializeComponent();
        }

        private void frmAnaDivForTest_Load(object sender, EventArgs e)
        {
            FillData();
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }
        private void frmAnaDivForTest_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Escape)
                Esci();
        }
        private void Esci()
        {
            this.Close();
        }

        private void FillData()
        {
            string s = "";
            string sTPath = ((string)_rowTes["trt_pth"]).Trim();
            string sTFile = ((string)_rowTes["trt_fil"]).Trim();
            string sTRego = ((string)_rowTes["trt_reg"]).Trim();

            string sFil = Path.GetFileName(sTPath);

            string[] sFils = Directory.GetFiles(Path.GetDirectoryName(sTPath), sTFile);

            foreach (string sFi in sFils)
            {
                s = Path.GetFileName(sFi).Substring(0,sFil.Length); 
                if(s == sFil)
                {
                    sFil = sFi;
                    break;
                }
            }

            DataTable t = new DataTable("tabTmp");

            DataView v = new DataView(_tabTab, "tra_des<>'' AND tra_ann=0", "tra_fld", DataViewRowState.CurrentRows);

            foreach (DataRowView r in v)
            {
                t.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = (string)r["tra_fld"],
                    Caption = (string)r["tra_des"],
                    MaxLength = Convert.ToInt16(r["tra_len"]),
                    ReadOnly = false
                });

            }

            string sRig = "";
            string sVal = "";
            string sFld = "";
            string sReg = "";       //Regola
            int iLen = 0;
            int iPos = 0;
            string sSep = "";
            int iRowIni = 0;

            if(sTRego != "")
            {
                string[] a = sTRego.Split(',');
                foreach(string ss in a)
                {
                    if (ss.Contains("SEPARATORE"))
                    {
                        string[] aa = ss.Trim().Split('-');
                        if (aa[1] == "PUNTOVIRGOLA")
                            sSep = ";";
                    }
                    if (ss.Contains("ROWINI"))
                    {
                        string[] aa = ss.Trim().Split('-');
                        if (_clsFun.Numerico(aa[1], "0123456789"))
                            iRowIni = Convert.ToInt16(aa[1]);
                    }
                }
            }

            using (StreamReader sr = new StreamReader(sFil, System.Text.Encoding.Default))
            {
                string[] aTes = sRig.Split('|');
                DataTable tTab = new DataTable();
                string sTab = "";
                Boolean bOk = true;
                int i = 0;

                while ((sRig = sr.ReadLine()) != null)
                {
                    i++;

                    if (sRig.Length >= 10 && i >= iRowIni)
                    {
                        if (Path.GetExtension(sTFile).ToLower() == ".csv")
                        {
                            bOk = true;
                            string[] aFie = sRig.Split(';');

                            DataRow x = t.NewRow();

                            foreach (DataRowView r in v)
                            {
                                sFld = (string)r["tra_fld"];

                                sReg = "";
                                if (!DBNull.Value.Equals(r["tra_reg"]))
                                    sReg = (string)r["tra_reg"];

                                iPos = Convert.ToInt16(r["tra_pos"]);

                                sVal = aFie[iPos - 1];
                                x[sFld] = sVal;
                            }

                            if (bOk)
                                t.Rows.Add(x);
                        }
                        else
                        {
                            bOk = true;
                            DataRow x = t.NewRow();

                            foreach (DataRowView r in v)
                            {
                                sFld = (string)r["tra_fld"];

                                sReg = "";
                                if (!DBNull.Value.Equals(r["tra_reg"]))
                                    sReg = (string)r["tra_reg"];

                                if (sFld == "div_qta")
                                    Console.WriteLine("aaa");

                                iLen = Convert.ToInt16(r["tra_len"]);
                                iPos = Convert.ToInt16(r["tra_pos"]);

                                if (sRig.Length >= iPos + iLen)
                                {
                                    sVal = sRig.Substring(iPos, iLen);
                                    x[sFld] = sVal;
                                }

                                if (sReg.Length > 1 && sReg.Substring(0, 2) == "==")
                                {
                                    if (sVal != sReg.Substring(2))
                                        bOk = false;
                                }
                            }
                            if(bOk)
                                t.Rows.Add(x);
                        }
                    }
                }
            }

            dgv1.DataSource = t;
        }

    }
}
