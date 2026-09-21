using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Threading;

namespace APOffice
{
    public partial class frmGesStatMovLegami : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        public string _strTip = "";
        public string _strArt = "";
        public string _strArd = "";
        public DataRow _rowLot;

        private string _strConSql = "";
        private string _strConSqlStat = "";

        public frmGesStatMovLegami()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
            _strConSql = _clsFun.ConSql(""); 
            _strConSqlStat = _clsFun.ConSql("3");
        }

        private void frmGesStaMovLegami_Load(object sender, EventArgs e)
        {
            SetDgv1();
            FillDati();
        }
        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }
        private void frmGesStaMovLegami_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Escape)
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
            DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "leg_rig";
            cTbc.Name = "Riga";
            cTbc.Width = 20;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "leg_leg";
            cTbc.Name = "Articolo";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_des";
            cTbc.Name = "Descrizione";
            cTbc.Width = 300;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.MaxInputLength = 90;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "leg_inc";
            cTbc.Name = "Inc. %";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 90;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "BilQta";
            cTbc.Name = "Q.tà bilancia";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 90;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "BilQkg";
            cTbc.Name = "Peso Kg bilancia";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 90;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "VenQta";
            cTbc.Name = "Q.tà movimentata";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 90;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "VenQkg";
            cTbc.Name = "Peso Kg movimentato";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 90;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns.Add(cTbc);

            cCbc = new DataGridViewCheckBoxColumn();
            cCbc.DataPropertyName = "leg_ann";
            cCbc.Name = "Ann.";
            cCbc.Width = 30;
            cCbc.ValueType = typeof(string);
            cCbc.ReadOnly = false;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cCbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "LegMdy";
            cTbc.Name = "Modificato";
            cTbc.Width = 0;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.Visible = false;
            dgv1.Columns.Add(cTbc);
        }

        private void FillDati()
        {
            string s = "";
            DataRow[] j;
            string sCau = "SCO";
            string sNeg = "001";
            DateTime dDdo = (DateTime)_rowLot["lot_ddo"];

            string sLot = (string)_rowLot["lot_cod"];
            string sYea = (string)_rowLot["lot_yea"];

            DataTable tTmp = new DataTable();

            s = "SELECT art_cod, art_reb, art_plu FROM AnaArticoli WHERE art_reb='1' AND art_plu <>''";
            DataTable tArt = _clsFun.FillTabSql("AnaArticoli", s, false, _strConSql);
            DataColumn[] Key = new DataColumn[2] 
            { 
                tArt.Columns["art_reb"],
                tArt.Columns["art_plu"] 
            };
            tArt.PrimaryKey = Key;

            if (_strArd == "")
            {
                s = "SELECT art_des FROM AnaArticoli WHERE art_cod='" + _strArt + "'";
                tTmp = _clsFun.FillTabSql("", s, false, _strConSql);
                if (tTmp.Rows.Count > 0)
                    _strArd = (string)tTmp.Rows[0]["art_des"];
            }

            tTmp = _clsQry.LotCarico(_rowLot);

            if (tTmp.Rows.Count > 0)
            {
                sNeg = (string)tTmp.Rows[0]["DocNeg"];
                dDdo = (DateTime)tTmp.Rows[0]["DocDdo"];
                lblQkg.Text = ((decimal)tTmp.Rows[0]["mov_qkg"]).ToString("#0.00");

                if (sNeg == "")
                    sNeg = "001";
            }

            lblArt.Text = _strArt + " " + _strArd + " Lotto: " + sLot;

            s = "SELECT AnaArtLegami.*, AnaArticoli.art_des FROM AnaArtLegami ";
            s += "LEFT OUTER JOIN AnaArticoli ON AnaArticoli.art_cod = AnaArtLegami.leg_leg ";
            s += "WHERE leg_tip='" + _strTip + "' AND leg_art='" + _strArt + "'";
            DataTable t = _clsFun.FillTabSql("", s, false, _strConSql);

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "LegMdy",
                Caption = "Mdy",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "BilQta",
                Caption = "Qta",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "BilQkg",
                Caption = "Peso",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "VenQta",
                Caption = "Qta",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "VenQkg",
                Caption = "Peso",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });

            s = _clsFun.ParGet(clsDefine.enuParametri.Par031Lotti2Pos, _strConSql);
            string sPrf = "";
            string[] a = s.Split(',');
            if (a.Length > 1 && a[1] != "")
                sPrf = a[1];

            sPrf = sPrf + Convert.ToInt16(sLot).ToString("000");

            s = "SELECT ven_art, ven_qta, ven_qkg FROM GesNegVen WHERE ";
            s += "ven_neg = '" + sNeg + "' AND ";
            s += "ven_cau = '" + sCau + "' AND ";
            s += "ven_day >= " + _clsFun.DaySql(dDdo) + "  AND ";
            s += "LEFT(ven_ean, 5) = '" + sPrf + "'";
            tTmp = _clsFun.FillTabSql("", s, false, _strConSqlStat);

            decimal dQkg = 0;

            foreach(DataRow y in tTmp.Rows)
            {
                j = t.Select("leg_leg='" + y["ven_art"] + "'");
                if(j.Length > 0)
                {
                    j[0]["VenQta"] = (decimal)j[0]["VenQta"] + (decimal)y["ven_qta"];
                    j[0]["VenQkg"] = (decimal)j[0]["VenQkg"] + (decimal)y["ven_qkg"];

                    dQkg += (decimal)y["ven_qkg"];
                }
            }

            s = "SELECT ";
            s += "GesMovTestate.mot_ddo AS DocDdo, ";
            s += "GesMovTestate.mot_neg AS DocNeg, ";
            s += "GesMovimenti.mov_art, ";
            s += "GesMovimenti.mov_qta, ";
            s += "GesMovimenti.mov_qkg, ";
            s += "TabMovCausali.tab_sgm ";
            s += "FROM GesMovTestate LEFT OUTER JOIN ";
            s += "GesMovimenti ON GesMovTestate.mot_ymo = GesMovimenti.mov_ymo AND GesMovTestate.mot_nmo = GesMovimenti.mov_nmo ";
            s += "LEFT OUTER JOIN TabMovCausali ON GesMovTestate.mot_cau = TabMovCausali.tab_cod ";
            s += "WHERE ";
            s += "tab_cfo <> '" + _clsDef.TIPFOR + "'  AND ";
            s += "mot_neg ='" + sNeg + "'  AND ";
            s += "mov_day >=" + _clsFun.DaySql(dDdo) + "  AND ";
            s += "mov_lot='" + sLot + "' ";
            tTmp = _clsFun.FillTabSql("TabMov", s, false, _strConSql);

            foreach (DataRow y in tTmp.Rows)
            {
                j = t.Select("leg_leg='" + y["mov_art"] + "'");
                if (j.Length > 0)
                {
                    int iSgn = 1;
                    if ((string)y["tab_sgm"] == "-")
                        iSgn = -1;

                    j[0]["VenQta"] = ((decimal)j[0]["VenQta"] + (decimal)y["mov_qta"]) * iSgn;
                    j[0]["VenQkg"] = ((decimal)j[0]["VenQkg"] + (decimal)y["mov_qkg"]) * iSgn; 

                    dQkg += (decimal)y["mov_qkg"];
                }
            }

            lblLegQkg.Text = dQkg.ToString("#0.00");
            dQkg = 0;

            string ss = sYea.Substring(2, 2) + sLot.Substring(1);

            s = "SELECT * FROM GesMovBilLotti WHERE ";
            s += "mol_day >=" + _clsFun.DaySql(dDdo) + "  AND ";
            s += "mol_lot = '" + ss + "'";
            tTmp = _clsFun.FillTabSql("TabMov", s, false, _strConSql);

            foreach (DataRow y in tTmp.Rows)
            {
                string sReb = ((string)y["mol_plu"]).Substring(1,1);
                string sPlu = Convert.ToInt32( ((string)y["mol_plu"]).Substring(2)).ToString();

                j = tArt.Select("art_reb='" + sReb + "' AND art_plu='" + sPlu + "'");
                if (j.Length > 0)
                {
                    string sArt = (string)j[0]["art_cod"];

                    j = t.Select("leg_leg='" + sArt + "'");
                    if (j.Length > 0)
                    {
                        j[0]["BilQta"] = (decimal)j[0]["BilQta"] + (decimal)y["mol_qta"];
                        j[0]["BilQkg"] = (decimal)j[0]["BilQkg"] + (decimal)y["mol_qkg"];

                        dQkg += (decimal)y["mol_qkg"];
                    }
                }
            }

            lblBilQkg.Text = dQkg.ToString("#0.00");

            dgv1.DataSource = t;
        }

        private void letturaBilanceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string s = "";
            string sPath = "C:\\ApProject\\Temp\\DivBilance\\";
            string sFile = "WinMori";

            Process p = new Process();
            p.StartInfo.FileName = "C:\\ApProject\\Temp\\DivBilance\\AP_BIZWINMORI.BAT";
            p.StartInfo.Arguments = "-r";
            p.StartInfo.ErrorDialog = true;
            p.StartInfo.WindowStyle = ProcessWindowStyle.Minimized;
            p.Start();
            p.WaitForExit(1000 * 60 * 7);    // wait up to 1 minutes.
            Thread.Sleep(4000);
            for (int i = 0; i <= 10; i++)
            {
                Thread.Sleep(4000);

                foreach (string sFil in Directory.GetFiles(sPath, sFile + "*"))
                {
                    if (File.Exists(sFil))
                    {
                        FilBilLotti(sFil);

                        s = Path.GetDirectoryName(sFil) + "\\Old\\" + Path.GetFileName(sFil);

                        File.Move(sFil, s);

                        break;
                    }
                }
            }

        }

        private void FilBilLotti(string strFil)
        {
            string s = "";
            string sLotLot = (string)_rowLot["lot_cod"];
            DateTime dLotDay = (DateTime)_rowLot["lot_ddo"];

            //string sFil = "C:\\WinSwGx-NET\\Winmori.dat";

            if(!File.Exists(strFil))
                MessageBox.Show(strFil + " non presente!", "CONTROLLO BIZERBA", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                string sRig = "";

                FileInfo fI = new FileInfo(strFil);
                progressBar1.Value = 0;
                progressBar1.Maximum = (int)fI.Length;
                progressBar1.Minimum = 0;

                using (StreamReader sr = new StreamReader(strFil))
                {
                    while (sr.Peek() >= 0)
                    {
                        progressBar1.Increment(sRig.Length);
                        Application.DoEvents();

                        sRig = sr.ReadLine();

                        string sLot = sRig.Substring(0, 6);
                        string sPlu = sRig.Substring(20, 6);
                        string sQta = sRig.Substring(26, 10);
                        string sQkg = sRig.Substring(36, 10);
                        string sDay = sRig.Substring(46, 8);

                        s = sDay.Substring(4,4) + sDay.Substring(2,2) + sDay.Substring(0,2);

                        Console.WriteLine("xxx");

                        DateTime dDay = _clsFun.Str2Day(s);

                        string sYea = dDay.Year.ToString();

                        if (DateTime.Compare(dDay, dLotDay) >= 0)
                        {
                            decimal dQta = Convert.ToDecimal(sQta);
                            decimal dQkg = Convert.ToDecimal(sQkg) / 1000;

                            s = "SELECT * FROM GesMovBilLotti WHERE ";
                            s += "mol_day=" + _clsFun.DaySql(dDay) + " AND ";
                            s += "mol_lot='" + sLot + "' AND ";
                            s += "mol_plu='" + sPlu + "'";
                            DataTable t = _clsFun.FillTabSql("GesMovBilLotti", s, false, _strConSql);

                            DataRow x = t.NewRow();
                            x["mol_yea"] = sYea;
                            x["mol_lot"] = sLot;
                            x["mol_plu"] = sPlu;
                            x["mol_qta"] = dQta;
                            x["mol_qkg"] = dQkg;
                            x["mol_day"] = dDay;

                            if (t.Rows.Count > 0)
                            {
                                x["mol_idx"] = t.Rows[0]["mol_idx"];
                                s = _clsFun.SqlUpdRowIdx("GesMovBilLotti", t, t.Rows[0], x, null);
                            }
                            else
                                s = _clsFun.SqlInsertRow("GesMovBilLotti", t, x);

                            if (s != "")
                                _clsFun.SqlWrite(s, _strConSql);

                        }

                        Console.WriteLine("zzzzzzz");
                    }
                    sr.Close();
                    sr.Dispose();
                }
            }
        }
        
    }
}
