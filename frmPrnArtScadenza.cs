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
    public partial class frmPrnArtScadenza : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        private string _strConSql = "";
        private string _strForRif = "";

        public frmPrnArtScadenza()
        {
            InitializeComponent(); 
            new clsGesGraph().SetGraph(this, 0);
            _strConSql = _clsFun.ConSql("");
        }

        private void frmPrnArtScadenza_Load(object sender, EventArgs e)
        {
            _strForRif = _clsFun.ParGet(clsDefine.enuParametri.ParForDefault, _strConSql);

            if(_strForRif != "")
            {
                string[] a = _strForRif.Split(',');
                if(a.Length > 0)
                    _strForRif = a[0];
            }

            dtpFin.Value = dtpIni.Value.AddDays(15);

            SetDgv1();
            FillDati();
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }
        private void frmPrnArtScadenza_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
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
            //dgv1.VirtualMode = true;
            //dgv1.Dock = DockStyle.Fill;
            dgv1.AllowUserToAddRows = false;
            dgv1.ReadOnly = false;
            dgv1.AllowUserToDeleteRows = false;

            DataGridViewTextBoxColumn cTbc;
            //DataGridViewCheckBoxColumn cCbc;
            DataGridViewComboBoxColumn cCmb;
            DataGridViewButtonColumn cBtn;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_cod";
            cTbc.Name = "Codice";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_arf";
            cTbc.Name = "C.Fornitore";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_des";
            cTbc.Name = "Descrizione";
            cTbc.Width = 250;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            //cTbc = new DataGridViewTextBoxColumn();
            //cTbc.DataPropertyName = "art_sca";
            //cTbc.Name = "Scadenza";
            //cTbc.Width = 100;
            //cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            //cTbc.ReadOnly = true;
            //dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "art_sca";
            cTbc.Name = "D.scadenza";
            cTbc.Width = 100;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.DefaultCellStyle.Format = "dd/MM/yy";
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cBtn = new DataGridViewButtonColumn();
            cBtn.UseColumnTextForButtonValue = true;
            cBtn.FlatStyle = FlatStyle.System;
            cBtn.Text = "X";
            cBtn.Width = 20;
            cBtn.ValueType = typeof(string);
            cBtn.ReadOnly = false;
            cBtn.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cBtn.DefaultCellStyle.BackColor = Color.White;
            cBtn.DefaultCellStyle.ForeColor = Color.Red;
            dgv1.Columns.Add(cBtn);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "TmpMdy";
            cTbc.Name = "Mdy";
            cTbc.Width = 0;
            cTbc.ValueType = typeof(string);
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            cTbc.Visible = false;
            dgv1.Columns.Add(cTbc);
        }

        private void btnArt_Click(object sender, EventArgs e)
        {
            frmSeekArt f = new frmSeekArt();
            f.ShowDialog();
            if (f._tabArt != null && f._tabArt.Rows.Count > 0)
            {
                DataTable t = f._tabArt;
                if (t.Rows.Count > 0)
                {
                    string sArt = (string)t.Rows[0]["tmp_art"];
                    string sArd = (string)t.Rows[0]["tmp_ard"];

                    DataTable tSca = (DataTable)dgv1.DataSource;

                    DataRow[] j = tSca.Select("art_cod='" + sArt +"'");
                    if(j.Length > 0 )
                        MessageBox.Show("Articolo già presente!", "INSERIMENTO ARTICOLO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else
                    {
                        DataRow xx = tSca.NewRow();
                        xx["art_cod"] = sArt;
                        xx["art_des"] = sArd;
                        xx["tmp_arf"] = "";

                        string s = "SELECT lia_arf FROM GesLisAcquisto WHERE lia_for='" + _strForRif + "' AND lia_art='" + sArt + "'";
                        DataTable tLia = _clsFun.FillTabSql("TLia", s, true, _strConSql);
                        if(tLia.Rows.Count > 0)
                            xx["tmp_arf"] = tLia.Rows[0]["lia_arf"];

                        tSca.Rows.Add(xx);
                    }
                }
            }
        }

        private void FillDati()
        {
            DataRow[] j;
            string s = "";

            s = "SELECT lia_art, lia_arf ";
            s += "FROM GesLisAcquisto ";
            s += "WHERE (lia_for = '" + _strForRif + "') ";
            s += "GROUP BY lia_art, lia_arf";
            DataTable tArf = _clsFun.FillTabSql("TabLia", s, false, _strConSql);
            DataColumn[] keys = new DataColumn[2];
            keys[0] = tArf.Columns["lia_arf"];
            keys[1] = tArf.Columns["lia_art"];
            tArf.PrimaryKey = keys;

            s = "SELECT art_cod, art_des, art_sca FROM AnaArticoli WHERE (NOT (art_sca IS NULL)) AND (art_sca <> " + _clsFun.DaySql(_clsDef.DAYOUT) +  ")";
            DataTable t = _clsFun.FillTabSql("TabArt", s, false, _strConSql);
            keys = new DataColumn[2];
            keys[0] = t.Columns["art_cod"];
            t.PrimaryKey = keys;

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "tmp_arf",
                Caption = "Articolo del fornitore primario",
                MaxLength = 20,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "TmpMdy",
                Caption = "Mdy",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            keys = new DataColumn[1];
            keys[0] = t.Columns["art_cod"];
            t.PrimaryKey = keys;

            progressBar1.Value = 0;
            progressBar1.Maximum = t.Rows.Count;
            progressBar1.Minimum = 0;

            foreach(DataRow y in t.Rows)
            {
                progressBar1.Increment(1);
                System.Windows.Forms.Application.DoEvents();

                j = tArf.Select("lia_art='" + y["art_cod"] + "'");
                if(j.Length > 0)
                    y["tmp_arf"] = (string)j[0]["lia_arf"];

                if ((DateTime)y["art_sca"] == _clsDef.DAYOUT)
                    y["art_sca"] = DBNull.Value;

            }

            dgv1.DataSource = t;
        }

        private void dgv1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 4)
            {
                dgv1.Rows[e.RowIndex].Cells["D.Scadenza"].Value = DBNull.Value;
                dgv1.Rows[e.RowIndex].Cells["Mdy"].Value = "S";
            }
        }

        private void dgv1_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dgv1.Columns[e.ColumnIndex].Name == "Descrizione")
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    string s = ((string)x["art_cod"]).Trim();

                    if (s != "")
                    {
                        frmAnaArticolo f = new frmAnaArticolo();
                        f._strArtCod = s;
                        f.ShowDialog();
                    }
                }
            }
        }  
        private void dgv1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 3)
            {
                //if (dgv1.Rows[e.RowIndex].Cells["Codice"].Value.ToString() != "")
                //{
                //Initialized a new DateTimePicker Control   
                dtp = new DateTimePicker();

                //if (dgv1.Rows[e.RowIndex].Cells["D.scadenza"].Value != null)
                //    dtp.Value = (DateTime)dgv1.Rows[e.RowIndex].Cells["D.scadenza"].Value;

                //if (!DBNull.Value.Equals(dgv1.Rows[e.RowIndex].Cells["D.scadenza"].Value))
                //    dtp.Value = (DateTime)dgv1.Rows[e.RowIndex].Cells["D.scadenza"].Value;

                if (e.RowIndex >= 0)
                {
                    if (!DBNull.Value.Equals(dgv1.Rows[e.RowIndex].Cells["D.scadenza"].Value))
                        dtp.Value = (DateTime)dgv1.Rows[e.RowIndex].Cells["D.scadenza"].Value;

                }


                //Adding DateTimePicker control into DataGridView    
                dgv1.Controls.Add(dtp);

                // Setting the format (i.e. 2014-10-10)   
                dtp.Format = DateTimePickerFormat.Short;

                // It returns the retangular area that represents the Display area for a cell   
                Rectangle oRectangle = dgv1.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true);

                //Setting area for DateTimePicker Control   
                dtp.Size = new Size(oRectangle.Width, oRectangle.Height);

                // Setting Location   
                dtp.Location = new Point(oRectangle.X, oRectangle.Y);

                // An event attached to dateTimePicker Control which is fired when DateTimeControl is closed   
                dtp.CloseUp += new EventHandler(oDateTimePicker_CloseUp);

                // An event attached to dateTimePicker Control which is fired when any date is selected   
                dtp.TextChanged += new EventHandler(dateTimePicker_OnTextChange);

                // Now make it visible   
                dtp.Visible = true;

                //dgv1.Rows[e.RowIndex].Cells["Mdy"].Value = "S";

                if (e.RowIndex >= 0)
                {
                    dgv1.Rows[e.RowIndex].Cells["Mdy"].Value = "S";
                }

                //}
            }
        }
        private void dateTimePicker_OnTextChange(object sender, EventArgs e)
        {
            // Saving the 'Selected Date on Calendar' into DataGridView current cell   
            dgv1.CurrentCell.Value = dtp.Text.ToString();
        }
        void oDateTimePicker_CloseUp(object sender, EventArgs e)
        {
            // Hiding the control after use    
            dtp.Visible = false;
        }

        private void Salva()
        {
            string s = "";

            DataTable t = (DataTable)dgv1.DataSource;

            //DataView v = new DataView(t,"TmpMdy='S'", "", DataViewRowState.CurrentRows);
            //foreach(DataRowView r in v)
            //{
            //    s = "SELECT art_cod, art_sca FROM AnaArticoli WHERE art_cod='" + r["art_cod"] + "'";
            //    DataTable tArt = _clsFun.FillTabSql("Art", s, true, _strConSql);
            //    if (tArt.Rows.Count > 0)
            //    {
            //        if(tArt.Rows[0]["art_sca"] != r["art_sca"])
            //        {
            //            s = "UPDATE AnaArticoli SET art_sca = " + _clsFun.DaySql((DateTime)r["art_sca"]) + " WHERE art_cod='" + r["art_cod"] + "'";
            //            _clsFun.SqlWrite(s, _strConSql);
            //        }
            //    }
            //}

            foreach (DataRow y in t.Rows)
            {
                if((string)y["TmpMdy"] == "S")
                {
                    s = "SELECT art_cod, art_sca FROM AnaArticoli WHERE art_cod='" + y["art_cod"] + "'";
                    DataTable tArt = _clsFun.FillTabSql("Art", s, true, _strConSql);

                    if (tArt.Rows.Count > 0)
                    {
                        if (tArt.Rows[0]["art_sca"] != y["art_sca"])
                        {
                            //s = "UPDATE AnaArticoli SET art_sca = " + _clsFun.DaySql((DateTime)y["art_sca"]) + " WHERE art_cod='" + y["art_cod"] + "'";

                            if (DBNull.Value.Equals(y["art_sca"]))
                                s = "UPDATE AnaArticoli SET art_sca = NULL WHERE art_cod='" + y["art_cod"] + "'";
                            else
                                s = "UPDATE AnaArticoli SET art_sca = " + _clsFun.DaySql((DateTime)y["art_sca"]) + " WHERE art_cod='" + y["art_cod"] + "'";

                            _clsFun.SqlWrite(s, _strConSql);

                        }
                    }
                }
            }

        }

        private void btnPrn_Click(object sender, EventArgs e)
        {
            PrnScadenze();
        }

        private void PrnScadenze()
        {
            TimeSpan ts = dtpFin.Value - dtpIni.Value;
            int iDif = ts.Days;

            DataColumn c;
            DataRow x;

            DataTable t = (DataTable)dgv1.DataSource;

            DataTable tDay = t.Clone();

            DateTime dDay = dtpIni.Value;

            for(int i = 0; i < iDif; i++)
            {

                DateTime dd = dDay.AddDays(i);
                string sDay = dd.ToString("MM-dd");

                tDay.Columns.Add(new DataColumn()
                {
                    DataType = Type.GetType("System.String"),
                    ColumnName = sDay,
                    Caption = sDay,
                    MaxLength = 1,
                    ReadOnly = false,
                    DefaultValue = (String)""
                });
            }

            Console.WriteLine("aaaaaaaa");

            foreach (DataRow y in t.Rows)
            {
                if (!DBNull.Value.Equals(y["art_sca"]))
                {
                    DateTime dd = (DateTime)y["art_sca"];

                    ts = dtpFin.Value - dtpIni.Value;
                    int ii  = ts.Days;

                    if (dd >= dtpIni.Value && dd <= dtpFin.Value)
                    {
                        string sDay = ((DateTime)y["art_sca"]).ToString("MM-dd");

                        x = tDay.NewRow();
                        x["art_cod"] = y["art_cod"];
                        x["art_des"] = y["art_des"];
                        x["art_sca"] = y["art_sca"];
                        x["tmp_arf"] = y["tmp_arf"];
                        x[sDay] = "X";
                        tDay.Rows.Add(x);
                    }
                }
            }

            dgv2.DataSource = tDay;

        }

        private void importTerminalinoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ImpTerm();
        }

        private void ImpTerm()
        {
            string s = "";
            string sMsg = "";
            DataRow[] j;

            DataTable tTer = new clsGenTabTmp().TabTmpTerRil("TabTer");

            frmGesImpTerm f = new frmGesImpTerm();
            //f._strTip = "INV";
            f._strTip = "SCA";
            f._tabTer = tTer;
            f.ShowDialog();
            if (f._tabImp != null)
            {
                DataTable t = f._tabImp.Copy();
                DataTable tSca = (DataTable)dgv1.DataSource;
                string sArd = "";

                progressBar1.Value = 0;
                progressBar1.Maximum = t.Rows.Count;
                progressBar1.Minimum = 0;

                foreach (DataRow y in t.Rows)
                {
                    progressBar1.Increment(1);
                    Application.DoEvents();

                    sArd = "";

                    if (((string)y["ord_art"]).Trim() == "")
                    {
                        DataTable tTmp = _clsQry.ArtSeek((string)y["ord_ean"], "");
                        if (tTmp.Rows.Count > 0)
                        {
                            y["ord_art"] = ((string)tTmp.Rows[0]["tmp_art"]).PadLeft(7, Convert.ToChar("0"));
                            sArd = (string)tTmp.Rows[0]["tmp_ard"];

                            //if ((string)tTmp.Rows[0]["tmp_sta"] == _clsDef.STANOA)
                            //    _clsQry.ArtAttiva((string)tTmp.Rows[0]["tmp_art"]);
                        }
                    }
                    else
                    {
                        DataTable tTmp = _clsQry.ArtSeek((string)y["ord_art"], "");
                        if (tTmp.Rows.Count > 0)
                            sArd = (string)tTmp.Rows[0]["tmp_ard"];
                    }

                    if (((string)y["ord_art"]).Trim() == "")
                        sMsg += "Non trovato ean " + (string)y["ord_ean"] + _clsDef.CRLF;
                    else
                    {
                        //DataRow x = tSca.NewRow();
                        //x["var_tip"] = _clsDef.VARETI;
                        //x["var_num"] = _clsDef.COD03Z;
                        //x["var_art"] = ((string)y["ord_art"]).PadLeft(7, Convert.ToChar("0"));
                        //x["var_ori"] = "Da terminalino";
                        //x["var_dti"] = DateTime.Today;
                        //x["var_dtv"] = DateTime.Today;

                        //if (strTip == "ETI")
                        //    FillEti(dSet, tEti, x, null);
                        //else
                        //{
                        //    FillPos(dSet, x);
                        //    FillEti(dSet, tEti, x, null);
                        //}

                        string sArt = ((string)y["ord_art"]).PadLeft(7, Convert.ToChar("0"));

                        s = ((decimal)y["ord_qta"]).ToString();

                        string[] a = s.Split(',');

                        if (a.Length < 2)
                            sMsg += (string)y["ord_ean"] + " " + sArd + " con data errata " + s +"! " + _clsDef.CRLF;
                        else
                        {

                            int iYea = DateTime.Today.Year;
                            int iDay = Convert.ToInt16(a[0]);
                            int iMon = Convert.ToInt16(a[1]);

                            if (iMon == 0 || iMon > 12 || iDay == 0 || iDay > 31)
                                sMsg += (string)y["ord_ean"] + " " + sArd + " con data errata " + s +"! " + _clsDef.CRLF;
                            else
                            {
                                DateTime dDay = new DateTime(iYea, iMon, iDay);


                                j = tSca.Select("art_cod='" + sArt + "'");
                                if (j.Length > 0)
                                    j[0]["art_sca"] = dDay;
                                else
                                {

                                    DataRow x = tSca.NewRow();
                                    x["art_cod"] = sArt;
                                    x["art_des"] = sArd;
                                    x["art_sca"] = dDay;
                                    x["tmp_arf"] = "";

                                    s = "SELECT lia_arf FROM GesLisAcquisto WHERE lia_for='" + _strForRif + "' AND lia_art='" + x["art_cod"] + "'";
                                    DataTable tLia = _clsFun.FillTabSql("TLia", s, true, _strConSql);
                                    if (tLia.Rows.Count > 0)
                                        x["tmp_arf"] = tLia.Rows[0]["lia_arf"];


                                    tSca.Rows.Add(x);
                                }
                            }
                        }
                    }
                }

            }

            if(sMsg != "")
                MessageBox.Show(sMsg, "ERRORI IMPORT", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
     }
}
