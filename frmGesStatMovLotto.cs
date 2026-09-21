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
    public partial class frmGesStatMovLotto : Form
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

        public frmGesStatMovLotto()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
            _strConSql = _clsFun.ConSql(""); 
            _strConSqlStat = _clsFun.ConSql("3");
        }

        private void frmGesStaMovLegami_Load(object sender, EventArgs e)
        {
            SetDgv1();
            //SetDgv2();
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
            cTbc.DataPropertyName = "tab_des";
            cTbc.Name = "Causale";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tab_sgm";
            cTbc.Name = "Segno";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "DocNdo";
            cTbc.Name = "N.Documento";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "DocDdo";
            cTbc.Name = "Data";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_l1d";
            cTbc.Name = "Settore";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_l2d";
            cTbc.Name = "Famiglia";
            cTbc.Width = 100;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_l3d";
            cTbc.Name = "SottoFamiglia";
            cTbc.Width = 90;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "mov_art";
            cTbc.Name = "Articolo";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "mov_ard";
            cTbc.Name = "Descrizione";
            cTbc.Width = 200;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "mov_qta";
            cTbc.Name = "Pezzi";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            //cTbc.MaxInputLength = 90;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "mov_qkg";
            cTbc.Name = "Peso";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_tkg";
            cTbc.Name = "Peso totale";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_liv";
            cTbc.Name = "Livello";
            cTbc.Width = 25;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

        }

        //private void SetDgv2()
        //{
        //    dgv2.AutoGenerateColumns = false;
        //    //dgv2.VirtualMode = true;
        //    //dgv2.Dock = DockStyle.Fill;
        //    dgv2.AllowUserToAddRows = false;
        //    dgv2.ReadOnly = false;
        //    dgv2.AllowUserToDeleteRows = false;
        //    //dgv2.DisplayedRowCount() = true;

        //    DataGridViewTextBoxColumn cTbc;
        //    DataGridViewCheckBoxColumn cCbc;
        //    //DataGridViewComboBoxColumn cCmb;

        //    //cTbc = new DataGridViewTextBoxColumn();
        //    //cTbc.DataPropertyName = "tab_des";
        //    //cTbc.Name = "Causale";
        //    //cTbc.Width = 70;
        //    //cTbc.ValueType = typeof(string);
        //    //cTbc.ReadOnly = true;
        //    ////cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        //    //dgv2.Columns.Add(cTbc);

        //    //cTbc = new DataGridViewTextBoxColumn();
        //    //cTbc.DataPropertyName = "tab_sgm";
        //    //cTbc.Name = "Segno";
        //    //cTbc.Width = 40;
        //    //cTbc.ValueType = typeof(string);
        //    //cTbc.ReadOnly = false;
        //    ////cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
        //    //dgv2.Columns.Add(cTbc);

        //    //cTbc = new DataGridViewTextBoxColumn();
        //    //cTbc.DataPropertyName = "DocNdo";
        //    //cTbc.Name = "N.Documento";
        //    //cTbc.Width = 60;
        //    //cTbc.ValueType = typeof(string);
        //    //cTbc.ReadOnly = true;
        //    //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
        //    //dgv2.Columns.Add(cTbc);

        //    //cTbc = new DataGridViewTextBoxColumn();
        //    //cTbc.DataPropertyName = "DocDdo";
        //    //cTbc.Name = "Data";
        //    //cTbc.Width = 70;
        //    //cTbc.ValueType = typeof(string);
        //    //cTbc.ReadOnly = false;
        //    //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        //    //dgv2.Columns.Add(cTbc);

        //    cTbc = new DataGridViewTextBoxColumn();
        //    cTbc.DataPropertyName = "tmp_l1d";
        //    cTbc.Name = "Settore";
        //    cTbc.Width = 100;
        //    cTbc.ValueType = typeof(string);
        //    cTbc.ReadOnly = false;
        //    dgv2.Columns.Add(cTbc);

        //    cTbc = new DataGridViewTextBoxColumn();
        //    cTbc.DataPropertyName = "tmp_l2d";
        //    cTbc.Name = "Famiglia";
        //    cTbc.Width = 100;
        //    cTbc.ValueType = typeof(string);
        //    cTbc.ReadOnly = false;
        //    dgv2.Columns.Add(cTbc);

        //    cTbc = new DataGridViewTextBoxColumn();
        //    cTbc.DataPropertyName = "tmp_l3d";
        //    cTbc.Name = "SottoFamiglia";
        //    cTbc.Width = 100;
        //    cTbc.ValueType = typeof(string);
        //    cTbc.ReadOnly = false;
        //    dgv2.Columns.Add(cTbc);

        //    cTbc = new DataGridViewTextBoxColumn();
        //    cTbc.DataPropertyName = "mov_art";
        //    cTbc.Name = "Articolo";
        //    cTbc.Width = 60;
        //    cTbc.ValueType = typeof(string);
        //    cTbc.ReadOnly = false;
        //    dgv2.Columns.Add(cTbc);

        //    cTbc = new DataGridViewTextBoxColumn();
        //    cTbc.DataPropertyName = "mov_ard";
        //    cTbc.Name = "Descrizione";
        //    cTbc.Width = 240;
        //    cTbc.ValueType = typeof(string);
        //    cTbc.ReadOnly = false;
        //    dgv2.Columns.Add(cTbc);

        //    cTbc = new DataGridViewTextBoxColumn();
        //    cTbc.DataPropertyName = "mov_qta";
        //    cTbc.Name = "Pezzi";
        //    cTbc.Width = 40;
        //    cTbc.ValueType = typeof(string);
        //    cTbc.ReadOnly = false;
        //    //cTbc.MaxInputLength = 90;
        //    cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        //    dgv2.Columns.Add(cTbc);

        //    cTbc = new DataGridViewTextBoxColumn();
        //    cTbc.DataPropertyName = "mov_qkg";
        //    cTbc.Name = "Peso";
        //    cTbc.Width = 50;
        //    cTbc.ValueType = typeof(string);
        //    cTbc.ReadOnly = false;
        //    cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
        //    dgv2.Columns.Add(cTbc);

        //    //cTbc = new DataGridViewTextBoxColumn();
        //    //cTbc.DataPropertyName = "tmp_liv";
        //    //cTbc.Name = "Livello";
        //    //cTbc.Width = 25;
        //    //cTbc.ValueType = typeof(string);
        //    //cTbc.ReadOnly = false;
        //    //dgv2.Columns.Add(cTbc);
        //}

        private void FillDatiOld()
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
                //lblQkg.Text = ((decimal)tTmp.Rows[0]["mov_qkg"]).ToString("#0.00");

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

            lblQkgSca.Text = dQkg.ToString("#0.00");
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

            //lblBilQkg.Text = dQkg.ToString("#0.00");

            dgv1.DataSource = t;
        }

        private void FillDati()
        {
            string s = "";
            DataRow x;
            DataRow[] j;

            s = "SELECT ";
            s += "AnaArticoli.art_cod,  ";
            s += "AnaArticoli.art_des,  ";
            //s += "AnaArticoli.art_sta,  ";
            //s += "AnaArticoli.art_iva,  ";
            //s += "AnaArticoli.art_umi,  ";
            //s += "AnaArticoli.art_tgr,  ";
            //s += "AnaArticoli.art_pxc,  ";
            //s += "AnaArticoli.art_pne,  ";
            //s += "AnaArticoli.art_rep,  ";
            s += "AnaArticoli.art_ec1,  ";
            s += "AnaArticoli.art_ec2,  ";
            s += "AnaArticoli.art_ec3,  ";
            //s += "AnaArticoli.art_reb,  ";
            //s += "AnaArticoli.art_plu,  ";
            //s += "AnaArticoli.art_bpz,  ";
            //s += "AnaArticoli.art_dti,  ";
            //s += "AnaArticoli.art_dtm,  ";
            //s += "AnaArticoli.art_bil,  ";
            //s += "AnaArticoli.art_cos,  ";
            //s += "AnaArticoli.art_prv,  ";
            s += "TabEcrLv1.tab_des AS EcrDe1,    ";
            s += "TabEcrLv2.tab_des AS EcrDe2,    ";
            s += "TabEcrLv3.tab_des AS EcrDe3,    ";
            s += "TabEcrLv3.tab_liv AS EcrLiv    ";
            //s += "TabReparti.tab_des AS RepDes,   ";
            //s += "TabStato.tab_des AS StaDes,     ";
            //s += "TabRepBilance.tab_des AS RebDes ";
            s += "FROM AnaArticoli               ";

            s += "LEFT OUTER JOIN TabEcrLv1 ON AnaArticoli.art_ec1 = TabEcrLv1.tab_cod ";
            s += "LEFT OUTER JOIN TabEcrLv2 ON AnaArticoli.art_ec1 = TabEcrLv2.tab_lv1 AND AnaArticoli.art_ec2 = TabEcrLv2.tab_cod ";
            s += "LEFT OUTER JOIN TabRepBilance ON AnaArticoli.art_reb = TabRepBilance.tab_cod ";
            //s += "LEFT OUTER JOIN TabStato ON AnaArticoli.art_sta = TabStato.tab_cod  ";
            s += "LEFT OUTER JOIN TabEcrLv3 ON AnaArticoli.art_ec1 = TabEcrLv3.tab_lv1 AND AnaArticoli.art_ec2 = TabEcrLv3.tab_lv2 AND AnaArticoli.art_ec3 = TabEcrLv3.tab_cod  ";
            //s += "LEFT OUTER JOIN TabReparti ON AnaArticoli.art_rep = TabReparti.tab_cod ";

            DataTable tArt = _clsFun.FillTabSql("AnaArticoli", s, false, _strConSql);

            DataColumn[] keys = new DataColumn[1];
            keys[0] = tArt.Columns["art_cod"];
            tArt.PrimaryKey = keys;

            DateTime dDdo = (DateTime)_rowLot["lot_ddo"];
            string sLot = (string)_rowLot["lot_cod"];
            string sYea = (string)_rowLot["lot_yea"];

            DataTable tTmp = _clsQry.LotCarico(_rowLot);

            if (tTmp.Rows.Count > 0)
            {
                //sNeg = (string)tTmp.Rows[0]["DocNeg"];
                //dDdo = (DateTime)tTmp.Rows[0]["DocDdo"];
                lblQkg.Text = ((decimal)tTmp.Rows[0]["mov_qkg"]).ToString("#0.00");

                //if (sNeg == "")
                //    sNeg = "001";
            }

            if (_strArd == "")
            {
                s = "SELECT art_des FROM AnaArticoli WHERE art_cod='" + _strArt + "'";
                tTmp = _clsFun.FillTabSql("", s, false, _strConSql);
                if (tTmp.Rows.Count > 0)
                    _strArd = (string)tTmp.Rows[0]["art_des"];
            }

            lblArt.Text = _strArt + " " + _strArd + " Lotto: " + sLot;
            
            s = "SELECT ";
            s += "TabDocTipo.tab_des, ";
            s += "TabDocTipo.tab_cod AS tab_sgm, ";
            s += "GesFatTestate.fat_ndo AS DocNdo, ";
            s += "GesFatTestate.fat_ddo AS DocDdo, ";
            s += "GesMovimenti.mov_art, ";
            s += "GesMovimenti.mov_ard, ";
            s += "GesMovimenti.mov_qta, ";
            s += "GesMovimenti.mov_qkg ";
            s += "FROM GesMovimenti ";
            s += "INNER JOIN GesFatTestate ON GesMovimenti.mov_yfa = GesFatTestate.fat_yfa AND GesMovimenti.mov_nfa = GesFatTestate.fat_nfa ";
            s += "LEFT OUTER JOIN TabDocTipo ON GesFatTestate.fat_tpd = TabDocTipo.tab_cod ";
            s += "WHERE ";
            s += "(GesMovimenti.mov_lot = 'XXXXX') AND ";
            s += "(GesMovimenti.mov_nmo = 'XXXXXX') AND ";
            s += "(TabDocTipo.tab_cod <> 'FA')";
            DataTable tFat = _clsFun.FillTabSql("GesMovimenti", s, false, _strConSql);

            s = "SELECT ";
            s += "TabMovCausali.tab_des, ";
            s += "TabMovCausali.tab_sgm, ";
            s += "TabMovCausali.tab_cfo, ";
            s += "TabMovCausali.tab_tip, ";
            s += "GesMovTestate.mot_ndo AS DocNdo, ";
            s += "GesMovTestate.mot_ddo AS DocDdo, ";
            s += "GesMovimenti.mov_art, ";
            s += "GesMovimenti.mov_ard, ";
            s += "GesMovimenti.mov_qta, ";
            s += "GesMovimenti.mov_qkg ";
            s += "FROM GesMovimenti ";
            s += "LEFT OUTER JOIN ";
            s += "GesMovTestate ON GesMovTestate.mot_ymo = GesMovimenti.mov_ymo AND ";
            s += "GesMovTestate.mot_nmo = GesMovimenti.mov_nmo ";
            s += "LEFT OUTER JOIN TabMovCausali ON GesMovTestate.mot_cau = TabMovCausali.tab_cod ";
            s += "WHERE ";
            s += "GesMovimenti.mov_nfa = 'XXXXXX' AND ";
            s += "GesMovimenti.mov_lot LIKE '%" + sLot + "%' ";
            DataTable tMov = _clsFun.FillTabSql("GesMovimenti", s, false, _strConSql);

            tFat.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_l1c",
                Caption = "Settore",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            tFat.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_l1d",
                Caption = "Settore descrizione",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            tFat.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_l2c",
                Caption = "Famiglia",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            tFat.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_l2d",
                Caption = "Fam descrizione",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            tFat.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_l3c",
                Caption = "SottoFamiglia",
                MaxLength = 3,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            tFat.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_l3d",
                Caption = "SottooFam descrizione",
                MaxLength = 50,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            tFat.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_liv",
                Caption = "Livello",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            tFat.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "tmp_tkg",
                Caption = "Peso porzionato e trasformato",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });

            decimal dQkgSez = 0;
            decimal dQkgCar = 0;
            decimal dQkgSca = 0;

            foreach(DataRow y in tMov.Rows)
            {
                if ((string)y["tab_cfo"] == "FOR" && ((string)y["tab_tip"]).Contains("D"))
                    Console.WriteLine("Non inserisco perchè documento di carico lotto");
                else
                    tFat.ImportRow(y);
            }

            foreach (DataRow y in tFat.Rows)
            {
                j = tArt.Select("art_cod='" + y["mov_art"] + "'");
                if(j.Length > 0)
                {
                    y["tmp_l1c"] = j[0]["art_ec1"];
                    y["tmp_l2c"] = j[0]["art_ec2"];
                    y["tmp_l3c"] = j[0]["art_ec3"];
                    y["tmp_l1d"] = j[0]["EcrDe1"];
                    y["tmp_l2d"] = j[0]["EcrDe2"];
                    y["tmp_l3d"] = j[0]["EcrDe3"];
                    y["tmp_liv"] = j[0]["EcrLiv"];

                    if (DBNull.Value.Equals(y["tmp_liv"]) || ((string)y["tmp_liv"]).Trim() == "")
                        y["tmp_liv"] = "3";
                }

                if ((string)y["tab_sgm"] == "FA")
                    y["tab_sgm"] = "+";
                else if ((string)y["tab_sgm"] == "NA")
                    y["tab_sgm"] = "-";

                if((string)y["tmp_liv"] == "2")
                    dQkgSez += (decimal)y["mov_qkg"];
                else if ((string)y["tmp_liv"] == "3")
                {
                    if ((string)y["tab_sgm"] == "+")
                        dQkgCar += (decimal)y["mov_qkg"];
                    else if ((string)y["tab_sgm"] == "-")
                        dQkgSca += (decimal)y["mov_qkg"];
                }

                if ((string)y["tmp_liv"] == "3")
                {
                    int iSgn = 1;
                    if ((string)y["tab_sgm"] == "-")
                        iSgn = -1;

                    decimal dPes = (decimal)y["mov_qkg"] * iSgn;

                    j = tFat.Select("tmp_l1c='" + y["tmp_l1c"] + "' AND tmp_l2c='" + y["tmp_l2c"] + "' AND tmp_liv='2'");
                    if (j.Length > 0)
                        j[0]["tmp_tkg"] = (decimal)y["tmp_tkg"] + dPes;
                }
            }

            //DataView v = new DataView(tFat, "", "", DataViewRowState.CurrentRows).ToTable();
            DataTable t = new DataView(tFat, "", "tmp_l1d,tmp_l2d, tmp_liv", DataViewRowState.CurrentRows).ToTable();

            dgv1.DataSource = t;

            lblQkgSez.Text = dQkgSez.ToString("#0.00");
            lblQkgCar.Text = dQkgCar.ToString("#0.00");
            lblQkgSca.Text = dQkgSca.ToString("#0.00");
            lblQkgTot.Text = (dQkgCar - dQkgSca).ToString("#0.00");

            //DataTable tRie = t.Clone();
            //foreach(DataRow y in t.Rows)
            //{
            //    if((string)y["tmp_liv"] == "3")
            //    {
            //        j = tRie.Select("mov_art='" + y["mov_art"] + "'");
            //        if(j.Length == 0)
            //        {
            //            x = tRie.NewRow();
            //            x["mov_art"] = y["mov_art"];
            //            x["mov_ard"] = y["mov_ard"];
            //            x["tmp_l1d"] = y["tmp_l1d"];
            //            x["tmp_l2d"] = y["tmp_l2d"];
            //            x["tmp_l3d"] = y["tmp_l3d"];
            //            x["mov_qta"] = 0;
            //            x["mov_qkg"] = 0;
            //            tRie.Rows.Add(x);
            //            j = tRie.Select("mov_art='" + y["mov_art"] + "'");
            //        }
            //        int iSgn = 1;
            //        if ((string)y["tab_sgm"] == "-")
            //            iSgn = -1;
            //        j[0]["mov_qta"] = (decimal)j[0]["mov_qta"] + ((decimal)y["mov_qta"] * iSgn) ;
            //        j[0]["mov_qkg"] = (decimal)j[0]["mov_qkg"] + ((decimal)y["mov_qkg"] * iSgn);
            //        decimal dPes = (decimal)y["mov_qkg"] * iSgn;
            //        j = t.Select("tmp_l1c='" + y["tmp_l1c"] + "' AND tmp_l2c='" + y["tmp_l2c"] + "' AND tmp_liv='2'");
            //        if (j.Length > 0)
            //            j[0]["tmp_tkg"] = (decimal)y["tmp_tkg"] + dPes;
            //    }
            //}
            //dgv2.DataSource = tRie;
        }

        private void dgv1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 8)
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    string sArt = (string)x["mov_art"];

                    frmAnaArticolo f = new frmAnaArticolo();
                    f._strArtCod = sArt;
                    f.ShowDialog();
                }
            }
        }

        private void excelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string s = "";
            decimal d = 0;
            DataRow[] j;

            //DataTable t = new DataTable();

            //DataView v = new DataView((DataTable)dgv1.DataSource, "", "DocDdo", DataViewRowState.CurrentRows);
            //t = v.Table.Clone();

            //foreach (DataRowView r in v)
            //{
            //    t.ImportRow(r.Row);
            //}

            DataTable t = ((DataTable)dgv1.DataSource).Copy();

            //DataRow x = t.NewRow();
            //x["tab_des"] = "Riepilogo porzionato";
            //t.Rows.Add(x);

            //v = new DataView((DataTable)dgv2.DataSource, "", "", DataViewRowState.CurrentRows);
            //foreach (DataRowView r in v)
            //{
            //    t.ImportRow(r.Row);
            //}

            DataRow x = t.NewRow();
            x["tab_des"] = "Totale sezionato";
            x["mov_qkg"] = Convert.ToDecimal(lblQkgSez.Text);
            t.Rows.Add(x);

            x = t.NewRow();
            x["tab_des"] = "Totale porzionato";
            x["mov_qkg"] = Convert.ToDecimal(lblQkgTot.Text);
            t.Rows.Add(x);


            //string sTit = "Giacenza - etrazione dati dal " + dtpIni.Value.ToString("dd/MM/yyyy") + " al " + DateTime.Now.ToString("dd/MM/yyyy");
            string sTit = this.Text + " al " + DateTime.Now.ToString("dd/MM/yyyy") + " - " + lblArt.Text + " Peso: " + lblQkg.Text;
            string sFoo = "";
            //sFoo += "'Totale peso a saldo',";
            //sFoo += ",,,,,," + lblQkgTot.Text.Replace(".", "").Replace(",", ".");
            string sFil = "";

            string sFld = "";
            sFld += "tab_des, Causale, 110, StringLiteral;";
            sFld += "tab_sgm, Segno, 40, StringLiteral;";
            sFld += "DocNdo, Documento, 45, StringLiteral;";
            sFld += "DocDdo, Data, 45, DateTime;";

            sFld += "tmp_l1d, Settore, 45, StringLiteral;";
            sFld += "tmp_l2d, Famiglia, 45, StringLiteral;";
            sFld += "tmp_l3d, SottoFamiglia, 45, StringLiteral;";

            sFld += "mov_art, Articolo, 45, StringLiteral;";
            sFld += "mov_ard, Descrizione, 45, StringLiteral;";
            sFld += "mov_qta, Colli, 45, Decimal;";
            sFld += "mov_qkg, Peso, 45, Decimal;";
            sFld += "tmp_tkg, Peso totale, 45, Decimal;";

            sFil = "LottoMovimenti.xls";

            (new clsExcel2()).exportToXls1(t, sFld, sFil, sTit, sFoo);
        }

        private void pDFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string sPar = lblArt.Text + "|";
            sPar += lblQkg.Text;
            
            new clsStampe().PrintMovLotto((DataTable)dgv1.DataSource, sPar);
        }

    }
}
