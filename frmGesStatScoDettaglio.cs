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
    public partial class frmGesStatScoDettaglio : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        public string _strNeg = "";
        public string _strCau = "";
        public DateTime _dayDay = new DateTime(2050, 1, 1, 0, 0, 0);
        public string _strOra = "";
        public string _strPos = "";
        public string _strSco = "";

        public DataTable _tabVet = new DataTable();

        private string _strConSql = "";
        private string _strConSqlSta = "";
        private string _strStatVisProf = "";

        public frmGesStatScoDettaglio()
        {
            InitializeComponent(); 
            new clsGesGraph().SetGraph(this, 0);
            _strConSql = _clsFun.ConSql("");
            _strConSqlSta = _clsFun.ConSql("3");
            _strStatVisProf = _clsFun.ParGet(clsDefine.enuParametri.ParStatVisProf, _strConSql);
        }

        private void frmGesStatScoDettaglio_Load(object sender, EventArgs e)
        {
            SetDgv1();
            SetDgv2();
            SetDgv3();
            FillDati();
        }

        private void frmGesStatScoDettaglio_KeyDown(object sender, KeyEventArgs e)
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
            if(modificaToolStripMenuItem.CheckState == CheckState.Checked)
                Salva();
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
            //DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vet_cau";
            cTbc.Name = "Causale";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            //cTbc = new DataGridViewTextBoxColumn();
            //cTbc.DataPropertyName = "ven_day";
            //cTbc.Name = "Data";
            //cTbc.Width = 50;
            //cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            //cTbc.ReadOnly = true;
            //dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vet_neg";
            cTbc.Name = "Negozio";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vet_pos";
            cTbc.Name = "Cassa";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vet_ora";
            cTbc.Name = "Ora";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vet_sco";
            cTbc.Name = "Scontrino";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vet_fid";
            cTbc.Name = "Tessera";
            cTbc.Width = 120;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vet_imp";
            cTbc.Name = "Importo";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vet_sct";
            cTbc.Name = "Sconti";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.MaxInputLength = 40;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns.Add(cTbc);

            //cTbc = new DataGridViewTextBoxColumn();
            //cTbc.DataPropertyName = "vet_scp";
            //cTbc.Name = "Sconto";
            //cTbc.Width = 50;
            //cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //cTbc.DefaultCellStyle.Format = "###,##0.00";
            //cTbc.ReadOnly = false;
            //dgv1.Columns.Add(cTbc);
            //cTbc = new DataGridViewTextBoxColumn();

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vet_pun";
            cTbc.Name = "Punti";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0";
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vet_scp";
            cTbc.Name = "Di cui punti articoli";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0";
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vet_sta";
            cTbc.Name = "Stato riga";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0";
            cTbc.ReadOnly = false;
            cTbc.ToolTipText = "R = importata in fatturazione";
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vet_doc";
            cTbc.Name = "Documento";
            cTbc.Width = 100;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
        }

        private void SetDgv2()
        {
            dgv2.AutoGenerateColumns = false;
            //dgv1.VirtualMode = true;
            //dgv1.Dock = DockStyle.Fill;
            dgv2.AllowUserToAddRows = false;
            dgv2.ReadOnly = true;
            dgv2.AllowUserToDeleteRows = false;
            //dgv1.DisplayedRowCount() = true;

            DataGridViewTextBoxColumn cTbc;
            //DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_ean";
            cTbc.Name = "Barcode";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_art";
            cTbc.Name = "Articolo";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_ard";
            cTbc.Name = "Descrizione";
            cTbc.Width = 160;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_iva";
            cTbc.Name = "IVA";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_umi";
            cTbc.Name = "UM";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_rep";
            cTbc.Name = "Reparto";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_qta";
            cTbc.Name = "Q.tà";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0";
            cTbc.ReadOnly = false;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_qkg";
            cTbc.Name = "Peso";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_prz";
            cTbc.Name = "Prezzo";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_sct";
            cTbc.Name = "Tipo riga";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = true;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_scn";
            cTbc.Name = "Sconto";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = true;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_ven";
            cTbc.Name = "Importo";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_pun";
            cTbc.Name = "Punti";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0";
            cTbc.ReadOnly = false;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_off";
            cTbc.Name = "Offerta";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = false;
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_sta";
            cTbc.Name = "Stato";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            //cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            cTbc.ToolTipText = "0=attivo, A=annullato";
            dgv2.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ven_ean";
            cTbc.Name = "Barcode";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = false;
            dgv2.Columns.Add(cTbc);
        }

        private void SetDgv3()
        {
            dgv3.AutoGenerateColumns = false;
            //dgv1.VirtualMode = true;
            //dgv1.Dock = DockStyle.Fill;
            dgv3.AllowUserToAddRows = false;
            dgv3.ReadOnly = true;
            dgv3.AllowUserToDeleteRows = false;
            //dgv1.DisplayedRowCount() = true;

            DataGridViewTextBoxColumn cTbc;
            //DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vep_tip";
            cTbc.Name = "Tipo riga";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv3.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "VepPag";
            cTbc.Name = "Tipo";
            cTbc.Width = 140;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv3.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vep_imp";
            cTbc.Name = "Importi";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv3.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "vep_val";
            cTbc.Name = "Valori";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv3.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "VepArt";
            cTbc.Name = "Articoli";
            cTbc.Width = 300;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv3.Columns.Add(cTbc);
        }

        private void FillDati()
        {
            DataRow[] j;

            string s = "SELECT * FROM GesNegVet WHERE ";
            s += "GesNegVet.vet_neg = '" + _strNeg + "' AND ";
            s += "GesNegVet.vet_cau = '" + _strCau + "' AND ";
            s += "GesNegVet.vet_day = " + _clsFun.DaySql(_dayDay) + " AND ";
            s += "GesNegVet.vet_ora = '" + _strOra + "' AND ";
            s += "GesNegVet.vet_pos = '" + _strPos + "' AND ";
            s += "GesNegVet.vet_sco = '" + _strSco + "' ";
            DataTable t = _clsFun.FillTabSql("Vet", s, false, _strConSqlSta);

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "VetMdy",
                Caption = "Mdy",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            dgv1.DataSource = t;

            s = "SELECT * FROM GesNegVen WHERE ";
            s += "GesNegVen.ven_neg = '" + _strNeg + "' AND ";
            s += "GesNegVen.ven_cau = '" + _strCau + "' AND ";
            s += "GesNegVen.ven_day = " + _clsFun.DaySql(_dayDay) + " AND ";
            s += "GesNegVen.ven_ora = '" + _strOra + "' AND ";
            s += "GesNegVen.ven_pos = '" + _strPos + "' AND ";
            s += "GesNegVen.ven_sco = '" + _strSco + "' ";
            t = _clsFun.FillTabSql("Ven", s, false, _strConSqlSta);

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "VenMdy",
                Caption = "Mdy",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            s = "SELECT * FROM TabCauCassa";
            DataTable tCau = _clsFun.FillTabSql("Cau", s, false, _strConSql);

            dgv2.DataSource = t;

            s = "SELECT * FROM GesNegVep WHERE ";
            s += "GesNegVep.vep_neg = '" + _strNeg + "' AND ";
            //s += "GesNegVep.vep_cau = '" + _strCau + "' AND ";
            s += "GesNegVep.vep_day = " + _clsFun.DaySql(_dayDay) + " AND ";
            s += "GesNegVep.vep_ora = '" + _strOra + "' AND ";
            s += "GesNegVep.vep_pos = '" + _strPos + "' AND ";
            s += "GesNegVep.vep_sco = '" + _strSco + "' ";
            t = _clsFun.FillTabSql("Vep", s, false, _strConSqlSta);
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "VepArt",
                Caption = "Articolo",
                MaxLength = 100,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "VepPag",
                Caption = "Pagamento",
                MaxLength = 100,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "VepMdy",
                Caption = "Mdy",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            foreach(DataRow y in t.Rows)
            {
                if (!DBNull.Value.Equals(y["vep_ean"]) && ((string)y["vep_ean"]).Trim() != "")
                {
                    DataTable tArt = _clsQry.SeekEanArt((string)y["vep_ean"]);
                    if (tArt.Rows.Count > 0)
                        y["VepArt"] = ((string)tArt.Rows[0]["tmp_ard"]).Trim() + "-" + (string)tArt.Rows[0]["tmp_art"];
                }

                j = tCau.Select("tab_cod='" + y["vep_cod"] + "'");
                if(j.Length > 0)
                    y["VepPag"] = (string)j[0]["tab_des"];
            }

            dgv3.DataSource = t;
        }

        private void modificaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (modificaToolStripMenuItem.CheckState == CheckState.Checked)
            {
                modificaToolStripMenuItem.CheckState = CheckState.Unchecked;
                //dgv2.Columns["ven_ven"].ReadOnly = true;
                //dgv3.Columns["vep_imp"].ReadOnly = true;
                dgv2.ReadOnly = true;
                dgv3.ReadOnly = true;
            }
            else
            {
                modificaToolStripMenuItem.CheckState = CheckState.Checked;
                //dgv2.Columns["ven_ven"].ReadOnly = false;
                //dgv3.Columns["vep_imp"].ReadOnly = false;
                dgv2.ReadOnly = false;
                dgv3.ReadOnly = false;
            }
        }

        private void dgv2_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            CurrencyManager cm = dgv2.BindingContext[dgv2.DataSource, dgv2.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {
                string sFld = dgv2.Columns[e.ColumnIndex].DataPropertyName;

                DataRowView r = cm.List[dgv2.CurrentRow.Index] as DataRowView;
                DataRow x = r.Row;
                
                if(sFld == "ven_ven")
                {
                    if ((string)x["ven_umi"] == "KG")
                        x["ven_prz"] = (decimal)x["ven_ven"] / (decimal)x["ven_qkg"];
                    else
                        x["ven_prz"] = (decimal)x["ven_ven"] / (decimal)x["ven_qta"];
                }
                
                x["VenMdy"] = "S";

                Calcola("VEN", "");
            }
        }

        private void dgv3_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {

            CurrencyManager cm = dgv3.BindingContext[dgv3.DataSource, dgv3.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {
                DataRowView r = cm.List[dgv3.CurrentRow.Index] as DataRowView;
                DataRow x = r.Row;

                string sTri = (string)x["vep_tip"];

                x["VepMdy"] = "S";
                Calcola("VEP", sTri);
            }
        }


        private void Calcola(string strTip, string  strTri)
        {
            DataTable tVet = (DataTable)dgv1.DataSource;
            DataTable tVen = (DataTable)dgv2.DataSource;
            DataTable tVep = (DataTable)dgv3.DataSource;

            if(strTip == "VEP" && strTri == "PUN")
            {
                decimal dPun = 0;

                foreach (DataRow y in tVen.Rows)
                {
                    if ((string)y["ven_sct"] != "PUN" && (decimal)y["ven_pun"] >0 )
                        dPun += (decimal)y["ven_pun"];
                }

                foreach (DataRow y in tVep.Rows)
                {
                    if ((string)y["vep_tip"] == "PUN" && (string)y["VepArt"] == "")
                        dPun += (decimal)y["vep_imp"];
                }

                if ((decimal)tVet.Rows[0]["vet_pun"] != dPun)
                {
                    tVet.Rows[0]["vet_pun"] = dPun;
                    tVet.Rows[0]["VetMdy"] = "S";
                }
            }

            if (strTip == "VEN"  || strTip == "VEP")
            {
                decimal dVen = 0;
                decimal dVep = 0;

                foreach (DataRow y in tVen.Rows)
                {
                    dVen += (decimal)y["ven_ven"];
                }

                foreach (DataRow y in tVep.Rows)
                {
                    if ((string)y["vep_tip"] == "PAG")
                        dVep += (decimal)y["vep_imp"];
                    if ((string)y["vep_tip"] == "RES")
                        dVep -= (decimal)y["vep_imp"];
                }

                if (dVep > 0 && dVep != dVen)
                    MessageBox.Show("Vendite scontrino (" + dVen.ToString("#0.00") + ") diverso dal totale pagamenti (" + dVep.ToString("#0.00") + ")", "CONTROLLO TOTALI", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                {
                    tVet.Rows[0]["vet_imp"] = dVen;
                    tVet.Rows[0]["VetMdy"] = "S";
                }
            }
        }

        private void Salva()
        {
            string s = "";
            DataRow[] j;
            DataRow x;

            ArrayList aWhe = new ArrayList();
            ArrayList aExl = new ArrayList();

            aWhe.Add("vet_neg");
            aWhe.Add("vet_cau");
            aWhe.Add("vet_day");
            aWhe.Add("vet_ora");
            aWhe.Add("vet_pos");
            aWhe.Add("vet_sco");

            s = "SELECT * FROM GesNegVet WHERE ";
            s += "GesNegVet.vet_neg = '" + _strNeg + "' AND ";
            s += "GesNegVet.vet_cau = '" + _strCau + "' AND ";
            s += "GesNegVet.vet_day = " + _clsFun.DaySql(_dayDay) + " AND ";
            s += "GesNegVet.vet_ora = '" + _strOra + "' AND ";
            s += "GesNegVet.vet_pos = '" + _strPos + "' AND ";
            s += "GesNegVet.vet_sco = '" + _strSco + "' ";
            DataTable t = _clsFun.FillTabSql("Vet", s, false, _strConSqlSta);

            DataTable tTmp = (DataTable)dgv1.DataSource;
            _tabVet = tTmp.Copy();

            s = _clsFun.SqlUpdRow("GesNegVet", t, t.Rows[0], tTmp.Rows[0], aWhe, aExl);
            if (s != "")
                _clsFun.SqlWrite(s, _strConSqlSta);

            aWhe.Clear();
            aWhe.Add("ven_neg");
            aWhe.Add("ven_cau");
            aWhe.Add("ven_day");
            aWhe.Add("ven_ora");
            aWhe.Add("ven_pos");
            aWhe.Add("ven_sco");

            s = "SELECT * FROM GesNegVen WHERE ";
            s += "GesNegVen.ven_neg = '" + _strNeg + "' AND ";
            s += "GesNegVen.ven_cau = '" + _strCau + "' AND ";
            s += "GesNegVen.ven_day = " + _clsFun.DaySql(_dayDay) + " AND ";
            s += "GesNegVen.ven_ora = '" + _strOra + "' AND ";
            s += "GesNegVen.ven_pos = '" + _strPos + "' AND ";
            s += "GesNegVen.ven_sco = '" + _strSco + "' ";
            t = _clsFun.FillTabSql("Ven", s, false, _strConSqlSta);

            tTmp = (DataTable)dgv2.DataSource;
            
            foreach(DataRow y in tTmp.Rows)
            {
                if ((string)y["VenMdy"] == "S")
                {
                    j = t.Select("ven_idx=" + Convert.ToString(y["ven_idx"]));
                    if (j.Length > 0)
                    {
                        s = _clsFun.SqlUpdRowIdx("GesNegVen", t, j[0], y, aExl);
                        if (s != "")
                            _clsFun.SqlWrite(s, _strConSqlSta);
                    }
                }
            }

            aWhe.Clear();
            aWhe.Add("vep_neg");
            aWhe.Add("vep_cau");
            aWhe.Add("vep_day");
            aWhe.Add("vep_ora");
            aWhe.Add("vep_pos");
            aWhe.Add("vep_sco");
            aWhe.Add("vep_cod");
            aWhe.Add("vep_tip");

            s = "SELECT * FROM GesNegVep WHERE ";
            s += "GesNegVep.vep_neg = '" + _strNeg + "' AND ";
            s += "GesNegVep.vep_cau = '" + _strCau + "' AND ";
            s += "GesNegVep.vep_day = " + _clsFun.DaySql(_dayDay) + " AND ";
            s += "GesNegVep.vep_ora = '" + _strOra + "' AND ";
            s += "GesNegVep.vep_pos = '" + _strPos + "' AND ";
            s += "GesNegVep.vep_sco = '" + _strSco + "' ";
            t = _clsFun.FillTabSql("Vep", s, false, _strConSqlSta);

            tTmp = (DataTable)dgv3.DataSource;

            foreach (DataRow y in tTmp.Rows)
            {
                if ((string)y["VepMdy"] == "S")
                {
                    //s = "";
                    //s += "vep_neg = '" + _strNeg + "' AND ";
                    //s += "vep_cau = '" + _strCau + "' AND ";
                    //s += "vep_day = " + _clsFun.DayMdb(_dayDay) + " AND ";
                    //s += "vep_ora = '" + _strOra + "' AND ";
                    //s += "vep_pos = '" + _strPos + "' AND ";
                    //s += "vep_sco = '" + _strSco + "' AND ";
                    //s += "vep_cod = '" + y["vep_cod"] + "' AND ";
                    //s += "vep_tip = '" + y["vep_tip"] + "'";
                    //j = t.Select(s);

                    j = t.Select("vep_sco='" + y["vep_sco"] + "' AND vep_cod='" + y["vep_cod"] + "' AND vep_ean='" + y["vep_ean"] + "' AND vep_fid='" + y["vep_fid"] + "'");
                    if(j.Length == 0)
                    {
                        s = _clsFun.SqlInsertRow("GesNegVep", t, y);
                        if (s != "")
                            _clsFun.SqlWrite(s, _strConSqlSta);
                    }
                    else
                    {
                        j = t.Select("vep_idx=" + Convert.ToString(y["vep_idx"]));
                        if (j.Length > 0)
                        {
                            s = _clsFun.SqlUpdRowIdx("GesNegVep", t, j[0], y, aExl);
                            if (s != "")
                                _clsFun.SqlWrite(s, _strConSqlSta);
                        }
                    }

                }
            }


        }

        private void dgv_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            e.CellStyle.BackColor = Color.Aquamarine;

            if (e.Control is TextBox)
            {
                var txt = e.Control as TextBox;
                if (txt != null)
                {
                    e.Control.KeyPress += new KeyPressEventHandler(txtPun2Virg_KeyPress);
                }
            }
        }

        private void txtPun2Virg_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((sender as TextBox).Text != null && (sender as TextBox).Text.Trim() != "")
            {
                if (e.KeyChar == '.')
                    e.KeyChar = ',';
            }
        }

        private void inserimentoPuntiFidelityToolStripMenuItem_Click(object sender, EventArgs e)
        {
            InsVep();
        }

        private void InsVep()
        {
            DataTable t = (DataTable)dgv1.DataSource;

            if (t.Rows.Count == 0 || DBNull.Value.Equals(t.Rows[0]["vet_fid"]) || ((string)t.Rows[0]["vet_fid"]) == "")
                MessageBox.Show("Tessera non definita nello scontrino!", "CONTROLLO PUNTI FIDELITY", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            else
            {
                string s = "";

                string sTes = (string)(string)t.Rows[0]["vet_fid"];

                s = "SELECT * FROM TabCauCassa WHERE tab_key='PUN'";
                DataTable tCau = _clsFun.FillTabSql("TabCauCassa", s, true, _strConSql);
                if (tCau.Rows.Count == 0)
                    MessageBox.Show("Causale cassa non definita per i punti fidelity!", "CONTROLLO PUNTI FIDELITY", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                {
                    string sCod = (string)tCau.Rows[0]["tab_cod"];
                    string sDes = (string)tCau.Rows[0]["tab_des"];
                    string sTip = (string)tCau.Rows[0]["tab_tip"];
                    string sTri = (string)tCau.Rows[0]["tab_tri"];

                    t = (DataTable)dgv3.DataSource;

                    DataView v = new DataView(t, "vep_tip='PUN'", "", DataViewRowState.CurrentRows);
                    if (v.Count > 0)
                        MessageBox.Show("Riga tipo punti già presente!", "CONTROLLO PUNTI FIDELITY", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    else
                    {
                        DataRow x = t.NewRow();
                        //x["vep_idx"] = 1;
                        x["vep_cau"] = t.Rows[0]["vep_cau"];
                        x["vep_neg"] = t.Rows[0]["vep_neg"];
                        x["vep_day"] = t.Rows[0]["vep_day"];
                        x["vep_ora"] = t.Rows[0]["vep_ora"];
                        x["vep_pos"] = t.Rows[0]["vep_pos"];
                        x["vep_sco"] = t.Rows[0]["vep_sco"];
                        x["VepPag"] = sDes;
                        x["vep_cod"] = sCod;
                        x["vep_tri"] = sTri;
                        x["vep_tip"] = sTip;
                        x["vep_off"] = "";
                        x["vep_imp"] = 0;
                        x["vep_val"] = 0;
                        x["vep_ppo"] = "";
                        x["vep_ord"] = "";
                        x["vep_ean"] = _clsDef.PUNTITESSERAX;
                        x["vep_fid"] = sTes;

                        t.Rows.Add(x);
                    }
                }
            }

        }

    }
}
