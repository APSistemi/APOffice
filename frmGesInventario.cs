using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System.IO;
using System.Diagnostics;

namespace APOffice
{
    public partial class frmGesInventario : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        private const string TABGESINT = "GesInvTestate";
        private const string TABGESINV = "GesInventario";
        private const string TABTABFIM = "TabForImport";

        public string _strInvNum = "";
        public string _strInvNeg = "";
        public string _strInvDes = "";

        private string _strConSql = "";

        public frmGesInventario()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
        }

        private void frmGesInventario_Load(object sender, EventArgs e)
        {
            ApplyModernUi();
            _strConSql = _clsFun.ConSql("");

            if (_strInvNeg == "")
                _strInvNeg = "001";

            lblIntNum.Text = _strInvNum;
            SetDgv1();
            FillTabs();
            FillDati();

            clsUiIcons.RestoreFormBounds(this);
            clsUiIcons.RestoreGridColumnWidths(dgv1, "frmGesInventario_dgv1");
        }

        private void frmGesInventario_FormClosing(object sender, FormClosingEventArgs e)
        {
            clsUiIcons.SaveFormBounds(this);
            clsUiIcons.SaveGridColumnWidths(dgv1, "frmGesInventario_dgv1");
        }

        private void ApplyModernUi()
        {
            try
            {
                if (menuStrip1 != null)
                {
                    menuStrip1.Renderer = clsUiIcons.GetModernMenuRenderer();
                    if (esciToolStripMenuItem != null)
                        esciToolStripMenuItem.Image = clsUiIcons.GetIcon("exit", 16);
                    if (importazioneDaTerminalinoToolStripMenuItem != null)
                        importazioneDaTerminalinoToolStripMenuItem.Image = clsUiIcons.GetIcon("barcode", 16);
                    if (utilitàToolStripMenuItem != null)
                        utilitàToolStripMenuItem.Image = clsUiIcons.GetIcon("tools_gear", 16);
                    if (simulatoreToolStripMenuItem != null)
                        simulatoreToolStripMenuItem.Image = clsUiIcons.GetIcon("settings", 16);
                    if (importAPPhoneToolStripMenuItem != null)
                        importAPPhoneToolStripMenuItem.Image = clsUiIcons.GetIcon("smartphone", 16);
                }

                this.BackColor = Color.FromArgb(243, 244, 246);

                if (btnPrn != null)
                {
                    clsUiIcons.StyleStatButton(btnPrn, "Stampa", "print", 14,
                        Color.FromArgb(254, 252, 232), Color.FromArgb(254, 240, 138),
                        Color.FromArgb(250, 204, 21), Color.FromArgb(113, 63, 18), Color.FromArgb(202, 138, 4));
                }

                if (btnPrnErr != null)
                {
                    clsUiIcons.StyleStatButton(btnPrnErr, "Errori", "warning", 14,
                        Color.FromArgb(254, 242, 242), Color.FromArgb(254, 202, 202),
                        Color.FromArgb(248, 113, 113), Color.FromArgb(127, 29, 29), Color.FromArgb(220, 38, 38));
                }

                if (btnXls != null)
                {
                    clsUiIcons.StyleStatButton(btnXls, "Excel", "excel", 14,
                        Color.FromArgb(236, 253, 245), Color.FromArgb(167, 243, 208),
                        Color.FromArgb(52, 211, 153), Color.FromArgb(6, 78, 59), Color.FromArgb(5, 150, 105));
                }

                if (btnAggCosti != null)
                {
                    clsUiIcons.StyleStatButton(btnAggCosti, "Agg. Costi", "refresh", 14,
                        Color.FromArgb(239, 246, 255), Color.FromArgb(191, 219, 254),
                        Color.FromArgb(96, 165, 250), Color.FromArgb(30, 58, 138), Color.FromArgb(29, 78, 216));
                }

                if (btnGia != null)
                {
                    clsUiIcons.StyleStatButton(btnGia, "Giacenze", "boxes", 14,
                        Color.FromArgb(245, 243, 255), Color.FromArgb(221, 214, 254),
                        Color.FromArgb(167, 139, 250), Color.FromArgb(76, 29, 149), Color.FromArgb(124, 58, 237));
                }

                if (btnDiff != null)
                {
                    clsUiIcons.StyleStatButton(btnDiff, "Differenze", "calculator", 14,
                        Color.FromArgb(240, 253, 250), Color.FromArgb(153, 246, 228),
                        Color.FromArgb(45, 212, 191), Color.FromArgb(19, 78, 74), Color.FromArgb(13, 148, 136));
                }

                if (dgv1 != null)
                {
                    clsUiIcons.StyleDataGridView(dgv1);
                }
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("frmGesInventario.ApplyModernUi", ex.Message);
            }
        }

        private void frmGesInventario_KeyDown(object sender, KeyEventArgs e)
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
            _strInvDes = txtIntDes.Text;
            Boolean b = Salva();
            if (!b)
            {
                if (MessageBox.Show("Abbandono delle modifiche, confermi?", this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                      == DialogResult.Yes)
                    this.Close();
            }
            else this.Close();
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
            cTbc.DataPropertyName = "inv_nte";
            cTbc.Name = "N.terminalino";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "inv_nrv";
            cTbc.Name = "N.rilevazione";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "inv_nri";
            cTbc.Name = "N.riga";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "InvL1d";
            cTbc.Name = "ECR 1";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "InvL2d";
            cTbc.Name = "ECR 2";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "InvL3d";
            cTbc.Name = "ECR 3";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "InvRed";
            cTbc.Name = "Reparto";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "inv_ean";
            cTbc.Name = "Barcode";
            cTbc.Width = 90;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "inv_art";
            cTbc.Name = "Articolo";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "InvArd";
            cTbc.Name = "Descrizione";
            cTbc.Width = 180;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 50;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "InvUmi";
            cTbc.Name = "UM";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 50;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "inv_gia";
            cTbc.Name = "Q.tà giacenza";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "inv_qta";
            cTbc.Name = "Q.tà contata";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "InvDif";
            cTbc.Name = "Differenza giacenza contato";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "inv_cos";
            cTbc.Name = "Costo";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.000";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "InvImp";
            cTbc.Name = "Importo";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.000";
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "inv_msg";
            cTbc.Name = "Messaggio";
            cTbc.Width = 200;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            //cTbc.MaxInputLength = 50;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "inv_dti";
            cTbc.Name = "Data inserimento";
            cTbc.Width = 100;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.MaxInputLength = 50;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "inv_dtt";
            cTbc.Name = "Data terminale";
            cTbc.Width = 120;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.MaxInputLength = 50;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);
        }

        private void FillTabs()
        {
            string p = "TabReparti";
            string s = "SELECT * FROM TabReparti ORDER BY tab_des";
            DataTable t = _clsFun.FillTabSql(p, s, false, _strConSql);
            DataRow x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  Tutti";
            t.Rows.InsertAt(x, 0);
            cmbIntRep.DataSource = t;
            cmbIntRep.DisplayMember = "tab_des";
            cmbIntRep.ValueMember = "tab_cod";
            cmbIntRep.SelectedValue = "";

            p = "TabNegozi";
            s = "SELECT * FROM TabNegozi ORDER BY tab_cod";
            t = _clsFun.FillTabSql(p, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  Tutti";
            t.Rows.InsertAt(x, 0);
            cmbIntNeg.DataSource = t;
            cmbIntNeg.DisplayMember = "tab_des";
            cmbIntNeg.ValueMember = "tab_cod";
            cmbIntNeg.SelectedValue = _strInvNeg;

            p = "TabNegozi";
            s = "SELECT * FROM TabStato WHERE tab_cod='XXX'";
            t = _clsFun.FillTabSql(p, s, false, _strConSql);
            x = t.NewRow();
            x["tab_cod"] = "P";
            x["tab_des"] = "Parziale";
            t.Rows.Add(x);
            x = t.NewRow();
            x["tab_cod"] = "T";
            x["tab_des"] = "Totale";
            t.Rows.Add(x);
            x = t.NewRow();
            x["tab_cod"] = "G";
            x["tab_des"] = "Solo giacenza";
            t.Rows.Add(x);
            cmbIntTip.DataSource = t;
            cmbIntTip.DisplayMember = "tab_des";
            cmbIntTip.ValueMember = "tab_cod";
            cmbIntTip.SelectedValue = "P";
        }

        private void FillDati()
        {
            DataTable tEcr = _clsQry.tabEcr();
            DataRow[] j;

            string s = "";
            s = "SELECT * FROM GesInvTestate WHERE ";
            s += "int_neg='" + cmbIntNeg.SelectedValue.ToString() + "' AND ";
            s += "int_num='" + lblIntNum.Text + "' ";
            DataTable t = _clsFun.FillTabSql(TABGESINT, s, false, _strConSql);
            if (t.Rows.Count > 0)
            {
                dtpIntDay.Value = (DateTime)t.Rows[0]["int_day"];
                txtIntDes.Text = (string)t.Rows[0]["int_des"];

                if(!DBNull.Value.Equals(t.Rows[0]["int_tip"]))
                {
                    //if ((string)t.Rows[0]["int_tip"] == "T")
                    //    cmbIntTip.Text = "Totale";
                    //else
                    //    cmbIntTip.Text = "Parziale";
                    cmbIntTip.SelectedValue = (string)t.Rows[0]["int_tip"];
                }

                if (!DBNull.Value.Equals(t.Rows[0]["int_rep"]))
                    cmbIntRep.SelectedValue = (string)t.Rows[0]["int_rep"];
            }

            s = "SELECT * FROM TabReparti";
            DataTable tRep = _clsFun.FillTabSql("TabRep", s, false, _strConSql);

            s = "SELECT ";
            s += "GesInventario.*, ";
            s += "AnaArticoli.art_des AS InvArd, ";
            s += "AnaArticoli.art_umi AS InvUmi, ";
            s += "AnaArticoli.art_iva AS InvIva, ";
            s += "AnaArticoli.art_ec1 AS InvL1c, ";
            s += "AnaArticoli.art_ec2 AS InvL2c, ";
            s += "AnaArticoli.art_ec3 AS InvL3c, ";
            s += "AnaArticoli.art_rep AS InvRep ";
            s += "FROM GesInventario ";
            s += "LEFT OUTER JOIN AnaArticoli ON GesInventario.inv_art = AnaArticoli.art_cod ";
            s += "WHERE ";
            s += "inv_neg='" + cmbIntNeg.SelectedValue.ToString() +"' AND ";
            s += "inv_num='" + lblIntNum.Text + "' ";
            s += "ORDER BY inv_nte, inv_nrv";

            t = _clsFun.FillTabSql(TABGESINV, s, false, _strConSql);
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "InvMdy",
                Caption = "Mdy",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "InvImp",
                Caption = "Imp",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "InvL1d",
                Caption = "Imp",
                MaxLength = 30,
                ReadOnly = false,
                DefaultValue = (string)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "InvL2d",
                Caption = "Imp",
                MaxLength = 30,
                ReadOnly = false,
                DefaultValue = (string)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "InvL3d",
                Caption = "Imp",
                MaxLength = 30,
                ReadOnly = false,
                DefaultValue = (string)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "InvRed",
                Caption = "Imp",
                MaxLength = 30,
                ReadOnly = false,
                DefaultValue = (string)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "InvDif",
                Caption = "Differenza",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "InvPrv",
                Caption = "Prezzo vendita",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });

            Boolean bGia = false;

            foreach (DataRow y in t.Rows)
            {
                if (DBNull.Value.Equals(y["inv_qta"]))
                    y["inv_msg"] = "";
                if (DBNull.Value.Equals(y["inv_qta"]))
                    y["inv_qta"] = 0;
                if (DBNull.Value.Equals(y["inv_cos"]))
                    y["inv_cos"] = 0;

                s = "l1c='" + y["InvL1c"] + "' AND l2c='" + y["InvL2c"] + "'AND l3c='" + y["InvL3c"] + "'";
                j = tEcr.Select(s);
                if (j.Length > 0)
                {
                    y["InvL1d"] = (string)j[0]["l1d"];
                    y["InvL2d"] = (string)j[0]["l2d"];
                    y["InvL3d"] = (string)j[0]["l3d"];
                }

                s = "tab_cod='" + y["InvRep"] + "'";
                j = tRep.Select(s);
                if (j.Length > 0)
                    y["InvRed"] = (string)j[0]["tab_des"];

                y["InvImp"] = (decimal)y["inv_qta"] * (decimal)y["inv_cos"];

                y["InvDif"] = (decimal)y["inv_qta"] - (decimal)y["inv_gia"];

                if ((decimal)y["inv_gia"] != 0)
                    bGia = true;
            }

            chkGia.Checked = bGia;

            DataView v = new DataView(t, "", "", DataViewRowState.CurrentRows);

            dgv1.DataSource = v;
            Totali();
            NewRiga(false);
        }

        private void NewRiga(Boolean bolReWrite)
        {
            Boolean b = true;

            DataTable t = ((DataView)dgv1.DataSource).Table;

            DataView v = new DataView(t, "inv_nri=''", "inv_art", DataViewRowState.CurrentRows);

            if (v.Count > 0)
            {
                b = false;

                if(bolReWrite)
                {
                    v[0].Delete();
                    b = true;
                }
            }

            if (b)
            {
                DataRow x = t.NewRow();
                x["inv_num"] = "";
                x["inv_nte"] = "00";
                x["inv_nrv"] = "00";
                x["inv_nri"] = "";
                x["inv_ean"] = "";
                x["inv_art"] = "";
                x["inv_gia"] = 0;
                x["inv_qta"] = 0;
                x["inv_ret"] = 0;
                x["inv_cos"] = 0;
                x["inv_msg"] = "";
                t.Rows.Add(x);

                CurrentCell("RIGNEW");
            }
        }

        private void FillInvRow(DataTable tabTmp, DataRow x, Boolean bolNew)
        {
            DataView v = new DataView(((DataView)dgv1.DataSource).ToTable(), "inv_art='" + (string)tabTmp.Rows[0]["tmp_art"] + "'", "", DataViewRowState.CurrentRows);
            if (v.Count > 0)
            {
                MessageBox.Show("Articolo già presente " + tabTmp.Rows[0]["tmp_art"] + " " + tabTmp.Rows[0]["tmp_ard"] + "!");
                if (!bolNew)
                    x.Delete();
            }
            else
            {
                x["inv_art"] = tabTmp.Rows[0]["tmp_art"];
                x["InvArd"] = tabTmp.Rows[0]["tmp_ard"];

                DataTable t = _clsQry.ArtSeek(((string)tabTmp.Rows[0]["tmp_art"]).Trim(), "");
                if (t.Rows.Count > 0)
                    x["InvUmi"] = t.Rows[0]["tmp_umi"];

                t = _clsQry.ArtCosto("", "", t, dtpIntDay.Value);
                if (t.Rows.Count > 0)
                    x["inv_cos"] = t.Rows[0]["tmp_cos"];
                
                x["InvImp"] = (decimal)x["inv_cos"] * (decimal)x["inv_qta"];
                x["InvMdy"] = "S";

                if (bolNew)
                {
                    int i = CurrentCell("");

                    if (i >= 0 && dgv1.Rows[i].Cells[0].Value.ToString() != "" && dgv1.Rows.Count - 1 == i)
                    {
                        NewRiga(false);
                    }
                }
            }
        }

        private void dgv1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 9 && e.RowIndex >= 0)
            {
                string sArt = dgv1.Rows[e.RowIndex].Cells["Articolo"].Value.ToString();

                if (sArt.Trim() == "")
                {
                    frmSeekArt f = new frmSeekArt();
                    f.ShowDialog();
                    if (f._tabArt != null && f._tabArt.Rows.Count > 0)
                    {
                        DataTable t = f._tabArt;
                        if (t.Rows.Count > 0)
                        {
                            CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                            if (cm.Position >= 0)
                            {
                                DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                                DataRow x = r.Row;

                                FillInvRow(t, x, true);
                            }
                        }
                    }
                }
                else
                {
                    frmAnaArticolo f = new frmAnaArticolo();
                    f._strArtCod = sArt;
                    f.ShowDialog();
                }
            }
        }
        private void dgv1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            TextBox txb = e.Control as TextBox;
            string s = "";

            if (txb != null)
            {
                e.Control.KeyPress += new KeyPressEventHandler(txt_KeyPress);
                txb.PreviewKeyDown += (S, E) =>
                {
                    if (E.KeyCode == Keys.Enter && s == "")
                    {
                        if (dgv1.CurrentCell.OwningColumn.Name.Equals("Descrizione"))
                        {
                            Console.WriteLine("aaaaaaa");
                            if (txb.Text != "")
                            {
                                DataTable t = _clsQry.ArtSeek(txb.Text, "SEEK");
                                if (t != null && t.Rows.Count > 0)
                                {
                                    CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                                    if (cm.Position >= 0)
                                    {
                                        DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                                        DataRow x = r.Row;
                                        FillInvRow(t, x, true);
                                        s = "Fatto";
                                    }
                                }
                            }
                        }
                        s = "Fatto";
                    }
                    return;
                };
            }
        }

        private void txt_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((sender as TextBox).Text != null && (sender as TextBox).Text.Trim() != "")
            {
                if (e.KeyChar == '.')
                    e.KeyChar = ',';
            }
        }

        private void dgv1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {
                DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                DataRow x = r.Row;
                //x["InvImp"] = (decimal)x["inv_cos"] * (decimal)x["inv_qta"];


                if (DBNull.Value.Equals(x["inv_cos"]))
                    x["inv_cos"] = 0;
                else if (DBNull.Value.Equals(x["inv_qta"]))
                    x["inv_qta"] = 0;
                if ((decimal)x["inv_qta"] > 0 && (decimal)x["inv_cos"] > 0)
                    x["InvImp"] = (decimal)x["inv_cos"] * (decimal)x["inv_qta"];

                if((string)x["inv_nri"] == "")
                {
                    CtrlNumRiga(x);
                }

                x["InvMdy"] = "S";
                CtrlNumRiga(x);
                cm.EndCurrentEdit();
                Totali();
                NewRiga(false);
            }
        }

        private void dgv1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dgv1.Columns[e.ColumnIndex].Name == "Ann." && dgv1.CurrentCell is DataGridViewCheckBoxCell)
            {
                dgv1.Rows[e.RowIndex].Cells["Modificato"].Value = "S";
            }
        }

        private void dgv1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dgv1.CurrentCell is DataGridViewCheckBoxCell)
            {
                dgv1.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }

        private void dgv1_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            dgv1.Rows[e.RowIndex].ErrorText = "";

            if (e.ColumnIndex == 11 || e.ColumnIndex == 12)
            {
                if (dgv1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value != null)
                {
                    string s = e.FormattedValue.ToString().Trim();

                    if (!_clsFun.Numerico(s, "0123456789,."))
                    {
                        e.Cancel = true;
                        dgv1.Rows[e.RowIndex].ErrorText = "Valore non corretto!";

                        MessageBox.Show("Dato inserito non corretto!", "RICHIESTO VALORE NUMERICO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void CtrlNumRiga(DataRow rowMov)
        {
            if ((string)rowMov["inv_nri"] == "" && ((string)rowMov["inv_art"]) != "")
            {
                DataTable t = ((DataView)dgv1.DataSource).ToTable();
                DataView v = new DataView(t, "inv_nri<>''", "inv_nri DESC", DataViewRowState.CurrentRows);
                int i = 0;

                if (v.Count > 0)
                    i = Convert.ToInt32(v[0]["inv_nri"]);
                rowMov["inv_nri"] = (i + 1).ToString("00000");
            }
        }

        private int CurrentCell(string strTip)
        {
            int iPos = 2;
            if (strTip == "RIGQTA")
                iPos = 4;
            int i = dgv1.Rows.Count - 1;
            if (i >= 0)
            {
                dgv1.Rows[i].Selected = true;
                dgv1.CurrentCell = dgv1[iPos, i];
            }
            return i;
        }

        private void Totali()
        {
            decimal dCos = 0;
            int iRig = 0;
            int iCol = 0;

            DataTable t = ((DataView)dgv1.DataSource).Table;


            foreach (DataRow y in t.Rows)
            {
                if (rdbPrnQtaRil.Checked)
                {
                    if (!DBNull.Value.Equals(y["inv_art"]) && (string)y["inv_art"] != "" && (decimal)y["inv_qta"] > 0)
                    {
                        iRig++;
                        iCol += Convert.ToInt32(y["inv_qta"]);

                        decimal dQta = 0;
                        decimal dPxc = 0;
                        decimal dPco = 0;
                        decimal dImp = 0;

                        if (!DBNull.Value.Equals(y["inv_qta"]))
                            dQta = Convert.ToDecimal(y["inv_qta"]);
                        if (!DBNull.Value.Equals(y["inv_cos"]))
                            dPco = Convert.ToDecimal(y["inv_cos"]);
                        dImp = dQta * dPco;

                        y["InvImp"] = dImp;

                        dCos += dImp;
                    }
                }
                else
                {
                    if (!DBNull.Value.Equals(y["inv_art"]) && (string)y["inv_art"] != "" && (decimal)y["inv_gia"] > 0)
                    {
                        iRig++;
                        iCol += Convert.ToInt32(y["inv_gia"]);

                        decimal dQta = 0;
                        decimal dPxc = 0;
                        decimal dPco = 0;
                        decimal dImp = 0;

                        if (!DBNull.Value.Equals(y["inv_gia"]))
                            dQta = Convert.ToDecimal(y["inv_gia"]);
                        if (!DBNull.Value.Equals(y["inv_cos"]))
                            dPco = Convert.ToDecimal(y["inv_cos"]);
                        dImp = dQta * dPco;

                        y["InvImp"] = dImp;

                        dCos += dImp;
                    }
                }
            }

            lblInvPez.Text = iCol.ToString("###,##0");
            lblInvRig.Text = iRig.ToString("###,##0");
            lblInvTot.Text = dCos.ToString("###,##0.00");
        }

        private Boolean InvCtrl()
        {
            Boolean b = true;

            DataView v = new DataView(((DataView)dgv1.DataSource).ToTable(), "inv_art <>'' AND (inv_gia > 0 OR inv_qta > 0)", "", DataViewRowState.CurrentRows);

            if (v.Count == 0)
            {
                b = false;
                MessageBox.Show("Non ci sono righe valide!");
            }

            return b;
        }

        private Boolean Salva()
        {
            string sParDivSede = _clsFun.ParGet(clsDefine.enuParametri.Par036Div2Sede, _strConSql);

            Boolean bOk = InvCtrl();
            string s = "";

            if (bOk)
            {
                Boolean b = false;
                DataTable t = new DataTable();
                ArrayList aWhe = new ArrayList();
                ArrayList aExl = new ArrayList();

                if(lblIntNum.Text == _clsDef.CODNEW)
                {
                    lblIntNum.Text = "01";
                    s = "SELECT * FROM " + TABGESINT + " ORDER BY int_num DESC";
                    t = _clsFun.FillTabSql(TABGESINT,s,true,_strConSql);
                    if(t.Rows.Count > 0)
                        lblIntNum.Text = (Convert.ToInt16(t.Rows[0]["int_num"]) +1).ToString("00");
                }

                s = "SELECT * FROM " + TABGESINT + " WHERE int_num='" + lblIntNum.Text + "'";
                t = _clsFun.FillTabSql(TABGESINT, s, false, _strConSql);

                DataRow y = t.NewRow();
                y["int_num"] = lblIntNum.Text;
                y["int_day"] = dtpIntDay.Value;
                y["int_des"] = txtIntDes.Text;
                y["int_neg"] = cmbIntNeg.SelectedValue.ToString();

                y["int_tip"] = "";
                if (cmbIntTip.Text.Length > 0)
                    y["int_tip"] = cmbIntTip.SelectedValue.ToString();

                y["int_rep"] = cmbIntRep.SelectedValue.ToString();

                if (t.Rows.Count == 0)
                {
                    s = _clsFun.SqlInsertRow(TABGESINT, t, y);
                }
                else
                {
                    aWhe = new ArrayList();
                    aExl = new ArrayList();
                    aWhe.Add("int_num");
                    s = _clsFun.SqlUpdRow(TABGESINT, t, t.Rows[0], y, aWhe, aExl);
                }
                if (s != "")
                {
                    b = true;
                    _clsFun.SqlWrite(s, _strConSql);
                    _clsFun.FileLog(TABGESINT, lblIntNum.Text, s);
                }

                s = "SELECT * FROM GesInventario WHERE ";
                s += "inv_num = '" + lblIntNum.Text + "' ";
                s += "ORDER BY inv_art";

                t = _clsFun.FillTabSql(TABGESINV, s, false, _strConSql);
                DataRow[] j;

                aWhe = new ArrayList();
                aWhe.Add("inv_num");
                aWhe.Add("inv_nte");
                aWhe.Add("inv_nrv");
                aWhe.Add("inv_nri");
                aWhe.Add("inv_art");

                aExl = new ArrayList();

                progressBar1.Value = 0;
                progressBar1.Maximum = (((DataView)dgv1.DataSource).ToTable()).Rows.Count;
                progressBar1.Minimum = 0;

                foreach (DataRow x in (((DataView)dgv1.DataSource).ToTable()).Rows)
                {
                    progressBar1.Increment(1);
                    System.Windows.Forms.Application.DoEvents();

                    if ((string)x["InvMdy"] == "S" && ((string)x["inv_art"]).Trim() != "")
                    {

                        x["inv_neg"] = cmbIntNeg.SelectedValue.ToString();

                        //s = "inv_num='" + x["inv_num"] + "' AND ";
                        //s += "inv_nte='" + x["inv_nte"] + "' AND ";
                        //s += "inv_nrv='" + x["inv_nrv"] + "' AND ";
                        //s += "inv_nri='" + x["inv_nri"] + "' AND ";
                        //s += "inv_art='" + x["inv_art"] + "'";

                        if(!DBNull.Value.Equals(x["inv_idx"]))
                        {
                            j = t.Select("inv_idx=" + Convert.ToString(x["inv_idx"]));
                            if (j.Length > 0)
                            s = _clsFun.SqlUpdRowIdx(TABGESINV, t, j[0], x, aExl);
                        }
                        else
                        {
                            x["inv_num"] = lblIntNum.Text;
                            s = _clsFun.SqlInsertRow(TABGESINV, t, x);
                        }

                        if (s != "")
                        {
                            b = true;
                            _clsFun.SqlWrite(s, _strConSql);
                            _clsFun.FileLog("GesInventario", (string)x["inv_art"], s);
                        }
                        x["InvMdy"] = "";
                    }
                }

                if(b && sParDivSede != "")
                {
                    string[] a = sParDivSede.Split(',');

                    if (a.Length > 1 && a[2] == "S")
                    {
                        string sNeg = _clsFun.FileIni("R", clsDefine.enuIni.Ini09CodiceAzienda, "");

                        clsVariazioni clsVar = new clsVariazioni();
                        string sPath = a[0];
                        string sPar = sPath + "," + lblIntNum.Text + "," + sNeg;

                        s = clsVar.DivNeg2SedeInventario(sPar);

                        if(s != "")
                            MessageBox.Show("Percorso non attivo!", "DIVULAGAZIONE A SEDE", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }

            return bOk;
        }
 
        private void importazioneDaTerminalinoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ImpTerm();
            if (dgv1.DataSource != null && ((DataView)dgv1.DataSource).Count > 0)
                InvPdfErr("ERR");
        }

        private void ImpTerm()
        {
            string s = "";
            DataRow[] j;
            Boolean b = true;

            string sNte = "01";
            string sNrv = "01";

            DataTable t = new DataTable();
            DataTable tEcr = _clsQry.tabEcr();
            DataTable tTer = new clsGenTabTmp().TabTmpTerRil("TabTer");
            DataTable tInv = ((DataView)dgv1.DataSource).Table;
            s = "SELECT * FROM TabReparti";
            DataTable tRep = _clsFun.FillTabSql("TabRep", s, false, _strConSql);

            if (tInv.Rows.Count > 0)
            {
                if ((string)tInv.Rows[tInv.Rows.Count - 1]["inv_art"] == "")
                    tInv.Rows[tInv.Rows.Count - 1].Delete();

                DataView v = new DataView(tInv,"inv_nte<>'00'","inv_nte, inv_nrv",DataViewRowState.CurrentRows);

                if(v.Count > 0)
                {
                    sNte = "";
                    sNrv = "";

                    foreach (DataRowView r in v)
                    {
                        if ((string)r["inv_nte"] != sNte || (string)r["inv_nrv"] != sNrv)
                        {
                            sNte = (string)r["inv_nte"];
                            if (_clsFun.Numerico(r["inv_nrv"],"0123456789"))
                                sNrv = (string)r["inv_nrv"];

                            j = tTer.Select("tmp_nte='" + sNte + "'");
                            if(j.Length == 0)
                            {
                                DataRow y = tTer.NewRow();
                                y["tmp_nte"] = (string)r["inv_nte"];
                                y["tmp_nrv"] = (string)r["inv_nrv"];
                                tTer.Rows.Add(y);
                            }
                            else
                            {
                                j[0]["tmp_nrv"] = (string)r["inv_nrv"];
                            }
                        }
                    }
                }
            }

            int i = tInv.Rows.Count;

            frmGesImpTerm f = new frmGesImpTerm();
            f._strTip = "INV";
            f._tabTer = tTer;
            f.ShowDialog();
            if (f._tabImp != null && f._tabImp.Rows.Count > 0)
            {
                string sMem = "";

                sNte = f._strDivTip.Substring(0, 2);
                sNrv = f._strDivTip.Substring(2, 2);

                foreach (DataRow y in f._tabImp.Rows)
                {
                    if ((string)y["ord_art"] == "0005197")
                        Console.WriteLine("aaaaa");
                    i++;
                    b = false;
                    //if (((string)y["ord_ean"]).Trim() != "")
                    if (_clsFun.Numerico(((string)y["ord_ean"]).Trim(),"0123456789"))
                    {
                        s = Convert.ToInt64(y["ord_ean"]).ToString();
                        t = _clsQry.ArtSeek(s, "");
                        b = true;
                    }
                    else if (((string)y["ord_art"]).Trim() != "")
                    {
                        s = Convert.ToInt64(y["ord_art"]).ToString();
                        t = _clsQry.ArtSeek(s, "");
                        b = true;
                    }

                    if(b)
                    {
                        if (t != null && t.Rows.Count > 0)
                        {
                            if (sMem == "")
                            {
                                s = (string)t.Rows[0]["tmp_art"];

                                j = tInv.Select("inv_nte='" + sNte + "' AND inv_nrv='" + sNrv + "' AND inv_art='" + s + "'");
                                if (j.Length > 0)
                                {
                                    s = (string)t.Rows[0]["tmp_art"] + " " + (string)t.Rows[0]["tmp_ard"] + _clsDef.CRLF;
                                    s += "ARTICOLO GIA' PRESENTE IN INVENTARIO, VUOI AGGIUNGERLO?";

                                    frmMsg1 f2 = new frmMsg1();
                                    f2._bolMem = true;
                                    f2._strTip = "SINO";
                                    f2._strMsg = s;
                                    f2.ShowDialog();
                                    s = f2._strRes;
                                    sMem = f2._strMem;

                                    if (s != "SI")
                                        b = false;
                                }
                            }
                            else if (sMem != "SI")
                            {
                                s = (string)t.Rows[0]["tmp_art"];
                                j = tInv.Select("inv_art='" + s + "'");
                                if (j.Length > 0)
                                    b = false;
                            }
                        }
                        if (b)
                        {
                            DataRow x = tInv.NewRow();
                            x["inv_num"] = lblIntNum.Text;
                            x["inv_nte"] = sNte;
                            x["inv_nrv"] = sNrv;
                            x["inv_nri"] = i.ToString("00000");
                            x["inv_ean"] = (string)y["ord_ean"];
                            x["inv_qta"] = (decimal)y["ord_qta"];
                            x["inv_msg"] = (string)y["ord_msg"];
                            x["inv_dti"] = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                            x["inv_dtt"] = (string)y["ord_dtt"];

                            if (((string)y["ord_msg"]).Trim() != "")
                                Console.WriteLine("aaaaaaaaaaaa");


                            x["inv_art"] = "";
                            x["InvArd"] = "";
                            x["InvL1c"] = "";
                            x["InvL2c"] = "";
                            x["InvL3c"] = "";
                            x["InvUmi"] = "";
                            x["inv_cos"] = 0;
                            //x["inv_msg"] = "";

                            if (t != null && t.Rows.Count > 0)
                            {
                                x["inv_art"] = (string)t.Rows[0]["tmp_art"];
                                x["InvArd"] = (string)t.Rows[0]["tmp_ard"];
                                x["InvUmi"] = (string)t.Rows[0]["tmp_umi"];
                                x["inv_cos"] = (decimal)t.Rows[0]["tmp_cos"];
                                x["inv_msg"] = "";

                                s = "l1c+l2c+l3c='" + (string)t.Rows[0]["tmp_ecr"] + "'";
                                j = tEcr.Select(s);
                                if (j.Length > 0)
                                {
                                    x["InvL1d"] = (string)j[0]["l1d"];
                                    x["InvL2d"] = (string)j[0]["l2d"];
                                    x["InvL3d"] = (string)j[0]["l3d"];
                                }

                                s = "tab_cod='" + (string)t.Rows[0]["tmp_rep"] + "'";
                                j = tRep.Select(s);

                                if (j.Length > 0)
                                    x["InvRed"] = (string)j[0]["tab_des"];
                            }
                            else
                            {
                                x["inv_msg"] = ((string)x["inv_msg"]).Trim() + " Barcode non trovato " + (string)y["ord_ean"];
                                if (((string)x["inv_msg"]).Length > 100)
                                    x["inv_msg"] = ((string)x["inv_msg"]).Substring(0, 100);

                            }
                            x["InvMdy"] = "S";
                            tInv.Rows.Add(x);
                        }
                    }
                }
                Totali();

                if (tInv.Rows.Count > 0)
                    NewRiga(false);
            }
        }

        private void btnPrnErr_Click(object sender, EventArgs e)
        {
            if (dgv1.DataSource == null || ((DataView)dgv1.DataSource).Count <= 1)
                MessageBox.Show("Non sono presenti articoli!", "CONTROLLO STAMPA");
            else
                InvPdfErr("ERR");
        }        

        private void btnPrn_Click(object sender, EventArgs e)
        {
            if (dgv1.DataSource == null || ((DataView)dgv1.DataSource).Count <= 1)
                MessageBox.Show("Non sono presenti articoli!", "CONTROLLO STAMPA");
            else
                InvPdf("ALL");
        }

        private void InvPdfErr(string strTip)
        {
            string s = "";
            DataRow[] j;
            int iRows = 37;
            int iRow = 0;

            DataView v = new DataView(((DataView)dgv1.DataSource).ToTable(), "", "inv_nri", DataViewRowState.CurrentRows);
            DataTable t = v.ToTable();

            DataTable tTmp = t.Clone();

            for (int i = 0; i < t.Rows.Count; i++ )
            {
                if (!DBNull.Value.Equals(t.Rows[i]["inv_msg"]) && ((string)t.Rows[i]["inv_msg"]).Trim() != "")
                {
                    tTmp.ImportRow(t.Rows[i]);
                    if (i > 0 && ((string)t.Rows[i-1]["inv_msg"]).Trim() == "")
                    {
                        tTmp.ImportRow(t.Rows[i-1]);
                    }
                    if (i < t.Rows.Count - 1 && ((string)t.Rows[i + 1]["inv_msg"]).Trim() == "")
                    {
                        tTmp.ImportRow(t.Rows[i + 1]);
                    }
                }
            }

            v = new DataView(tTmp, "", "inv_nri", DataViewRowState.CurrentRows);
            t = v.ToTable();

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Articoli";

            // Create an empty page
            PdfPage page = pd.AddPage();

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 10, XFontStyle.Bold);
            XFont font2 = new XFont("Courier new", 9, XFontStyle.Regular);
            XFont font3 = new XFont("Courier new", 8, XFontStyle.Regular);
            XFont font4 = new XFont("Courier new", 8, XFontStyle.Bold);

            double Y = 5;
            double X = 20;

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            string sFam = ""; // (string)t.Rows[0]["InvL1c"] + (string)t.Rows[0]["InvL2c"] + (string)t.Rows[0]["InvL3c"];

            decimal dRig = 0;
            decimal dPez = 0;
            decimal dVal = 0;

            for (int i = 0; i <= t.Rows.Count - 1; i++)
            {
                if (iRow > iRows || iRow == 0)
                {
                    if (iRow > 0)
                    {
                        page = pd.AddPage();
                        gfx = XGraphics.FromPdfPage(page);
                    }

                    X = 20;
                    iRow = 0;
                    s = DateTime.Now.ToString("dd/MM/yyyy") + " STAMPA INVENTARIO ERRORI AL " + dtpIntDay.Value.ToString("dd/MM/yyyy");
                    //if (chkCli.Checked)
                    //    s += "CLIENTI ";
                    //if (chkFor.Checked)
                    //    s += "FORNITORI ";
                    gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    X += 20;

                    s = "Codice";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 70, X, XStringFormats.Default);

                    s = "Descrizione";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 130, X, XStringFormats.Default);

                    s = "UM";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 300, X, XStringFormats.Default);

                    s = "Q.ta";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 340, X, XStringFormats.Default);

                    s = "Costo";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 400, X, XStringFormats.Default);

                    s = "Importo";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 450, X, XStringFormats.Default);

                    X += 15;
                }

                iRow++;

                if (t.Rows.Count - 1 >= i)
                {
                    //if (sFam != (string)t.Rows[i]["InvL1c"] + (string)t.Rows[i]["InvL2c"] + (string)t.Rows[i]["InvL3c"])
                    //{
                    //    if (i > 0)
                    //    {
                    //        s = "Famiglia totali righe " + dRig.ToString() + " ";
                    //        gfx.DrawString(s, font3, XBrushes.Black, Y + 10, X, XStringFormats.Default);
                    //        s = (dPez).ToString("###0.00");
                    //        gfx.DrawString(s, font3, XBrushes.Black, Y + 370, X, frmDX);
                    //        s = dVal.ToString("###,##0.00");
                    //        gfx.DrawString(s, font3, XBrushes.Black, Y + 480, X, frmDX);
                    //        X += 20;
                    //        iRow++;
                    //    }

                    //    dRig = 0;
                    //    dPez = 0;
                    //    dVal = 0;

                    //    sFam = (string)t.Rows[i]["InvL1c"] + (string)t.Rows[i]["InvL2c"] + (string)t.Rows[i]["InvL3c"];

                    //    s = "Famiglia " + (string)t.Rows[i]["InvL1c"] + " " + 
                    //        ((string)t.Rows[i]["InvL1d"]).Trim() + " - " + 
                    //        (string)t.Rows[i]["InvL2c"] + " " + 
                    //        ((string)t.Rows[i]["InvL2d"]).Trim() + " - " + 
                    //        (string)t.Rows[i]["InvL3c"] + " " + 
                    //        ((string)t.Rows[i]["InvL3d"]).Trim();
                    //    gfx.DrawString(s, font3, XBrushes.Black, Y + 10, X, XStringFormats.Default);
                    //    X += 20;
                    //    iRow++;
                    //}

                    s = "";
                    if (!DBNull.Value.Equals(t.Rows[i]["inv_ean"]))
                        s = (string)t.Rows[i]["inv_ean"];
                    gfx.DrawString(s, font4, XBrushes.Black, Y + 0, X, XStringFormats.Default);

                    if (!DBNull.Value.Equals(t.Rows[i]["inv_art"]))
                        s = (string)t.Rows[i]["inv_art"];
                    gfx.DrawString(s, font4, XBrushes.Black, Y + 70, X, XStringFormats.Default);

                    s = "";
                    if (!DBNull.Value.Equals(t.Rows[i]["InvArd"]))
                        s = (string)t.Rows[i]["InvArd"];
                    gfx.DrawString(s, font4, XBrushes.Black, Y + 130, X, XStringFormats.Default);

                    s = "";
                    if (!DBNull.Value.Equals(t.Rows[i]["InvUmi"]))
                        s = (string)t.Rows[i]["InvUmi"];
                    gfx.DrawString(s, font4, XBrushes.Black, Y + 300, X, XStringFormats.Default);

                    s = "";
                    if (!DBNull.Value.Equals(t.Rows[i]["inv_qta"]))
                        s = ((decimal)t.Rows[i]["inv_qta"]).ToString("####,##0.00");
                    gfx.DrawString(s, font4, XBrushes.Black, Y + 370, X + 3, frmDX);

                    s = "";
                    if (!DBNull.Value.Equals(t.Rows[i]["inv_cos"]))
                        s = ((decimal)t.Rows[i]["inv_cos"]).ToString("####,##0.000");
                    gfx.DrawString(s, font4, XBrushes.Black, Y + 430, X + 3, frmDX);

                    s = "";
                    if (!DBNull.Value.Equals(t.Rows[i]["InvImp"]))
                        s = ((decimal)t.Rows[i]["InvImp"]).ToString("####,##0.00");
                    gfx.DrawString(s, font4, XBrushes.Black, Y + 480, X + 3, frmDX);

                    //s = "";
                    //if (!DBNull.Value.Equals(t.Rows[i]["orf_msg"]))
                    //    s = (string)t.Rows[i]["orf_msg"];
                    //gfx.DrawString(s, font4, XBrushes.Black, Y + 30, X + 7, XStringFormats.Default);

                    //break;

                    dRig++;
                    dPez += (decimal)t.Rows[i]["inv_qta"];
                    dVal += (decimal)t.Rows[i]["InvImp"];

                }

                X += 20;
                //break;

            }

            //s = "Famiglia totali righe " + dRig.ToString() + " ";
            //gfx.DrawString(s, font3, XBrushes.Black, Y + 10, X, XStringFormats.Default);
            //s = (dPez).ToString("###0.00");
            //gfx.DrawString(s, font3, XBrushes.Black, Y + 370, X, frmDX);
            //s = dVal.ToString("###,##0.00");
            //gfx.DrawString(s, font3, XBrushes.Black, Y + 480, X, frmDX);

            //X += 25;

            s = "Totale righe " + lblInvRig.Text;
            s += "    Totale pezzi " + lblInvPez.Text;
            //s += "    Totale valore costo inventario   " + lblInvTot.Text;
            gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

            // Save the document...
            try
            {
                string sFil = "C:\\ApProject\\PDF\\InvErr_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                pd.Save(sFil);
                // ...and start a viewer.
                Process.Start(sFil);
            }
            catch (Exception ex)
            {
                MessageBox.Show("File di stampa già aperto!");
            }

        }

        private void InvPdf(string strTip)
        {
            string s = "";
            DataRow[] j;
            int iRows = 37;
            int iRow = 0;

            string sTer = "";           //Numero terminale
            string sRil = "";           //Numero rilevazione
            if (_clsFun.Numerico(txtRil.Text, "0123456789."))
            {
                string[] a = txtRil.Text.Split('.');
                if(a.Length > 1)
                {
                    sTer = a[0].PadLeft(2,Convert.ToChar('0'));
                    sRil = a[1].PadLeft(2, Convert.ToChar('0'));
                }
            }

            DataTable t = new DataTable();
            if (strTip == "ALL")
            {
                //t = ((DataView)dgv1.DataSource).ToTable();

                DataView v = new DataView(((DataView)dgv1.DataSource).ToTable(), "inv_msg ='' AND inv_art<>'' AND inv_qta>0 AND inv_cos=0", "InvL1d,InvL2d,InvL3d,inv_art", DataViewRowState.CurrentRows);
                //DataView v = new DataView(((DataView)dgv1.DataSource).ToTable(), "inv_msg ='' AND inv_art<>'' AND inv_qta>0", "InvL1d,InvL2d,InvL3d,inv_art", DataViewRowState.CurrentRows);

                if(v.Count > 0)
                    MessageBox.Show("Ci sono righe con costo a ZERO!", "STAMPA INVENTARIO", MessageBoxButtons.OK, MessageBoxIcon.Information);

                //Console.WriteLine("xxxxxxxxxx");
                //v = new DataView(((DataView)dgv1.DataSource).ToTable(), "inv_msg ='' AND inv_art<>'' AND inv_qta>0 AND inv_cos>0", "InvL1d,InvL2d,InvL3d,inv_art", DataViewRowState.CurrentRows);

                if (sRil != "")
                    s = " AND inv_nte='" + sTer + "' AND inv_nrv='" + sRil + "'";

                if (rdbPrnQtaGia.Checked)
                    v = new DataView(((DataView)dgv1.DataSource).ToTable(), "inv_msg ='' AND inv_art<>'' AND inv_gia<>0" + s, "InvL1d,InvL2d,InvL3d,inv_art", DataViewRowState.CurrentRows);
                else
                    v = new DataView(((DataView)dgv1.DataSource).ToTable(), "inv_msg ='' AND inv_art<>'' AND inv_qta<>0" + s, "InvL1d,InvL2d,InvL3d,inv_art", DataViewRowState.CurrentRows);

                t = v.Table.Clone();

                s = "";

                progressBar1.Value = 0;
                progressBar1.Maximum = v.Count;
                progressBar1.Minimum = 0;

                foreach (DataRowView r in v)
                {
                    if ((string)r["inv_art"] == "0000037")
                        Console.WriteLine("zzzzzz");

                    progressBar1.Increment(1);
                    System.Windows.Forms.Application.DoEvents();

                    j = t.Select("inv_art='" + r["inv_art"] + "'");
                    if (j.Length == 0)
                    {
                        DataRow x = t.NewRow();
                        x["inv_art"] = r["inv_art"];
                        x["InvArd"] = r["InvArd"];
                        x["InvL1c"] = r["InvL1c"];
                        x["InvL2c"] = r["InvL2c"];
                        x["InvL3c"] = r["InvL3c"];
                        x["InvL1d"] = r["InvL1d"];
                        x["InvL2d"] = r["InvL2d"];
                        x["InvL3d"] = r["InvL3d"];
                        x["InvUmi"] = r["InvUmi"];
                        x["inv_gia"] = 0;
                        x["inv_qta"] = 0;
                        x["inv_cos"] = 0;
                        x["InvImp"] = 0;
                        t.Rows.Add(x);
                        j = t.Select("inv_art='" + r["inv_art"] + "'");
                    }

                    if (rdbPrnQtaGia.Checked)
                    {
                        //j[0]["inv_gia"] = (decimal)j[0]["inv_gia"] + (decimal)r["inv_gia"];
                        //j[0]["InvImp"] = (decimal)j[0]["InvImp"] + ((decimal)r["inv_gia"] * (decimal)r["inv_cos"]);
                        //if ((decimal)j[0]["InvImp"] > 0 && (decimal)j[0]["inv_gia"] > 0)
                        //    j[0]["inv_cos"] = (decimal)j[0]["InvImp"] / (decimal)j[0]["inv_gia"];
                        j[0]["inv_qta"] = (decimal)j[0]["inv_qta"] + (decimal)r["inv_gia"];
                        j[0]["InvImp"] = (decimal)j[0]["InvImp"] + ((decimal)r["inv_gia"] * (decimal)r["inv_cos"]);
                        if ((decimal)j[0]["InvImp"] > 0 && (decimal)j[0]["inv_qta"] > 0)
                            j[0]["inv_cos"] = (decimal)j[0]["InvImp"] / (decimal)j[0]["inv_qta"];
                    }
                    else
                    {
                        j[0]["inv_qta"] = (decimal)j[0]["inv_qta"] + (decimal)r["inv_qta"];
                        j[0]["InvImp"] = (decimal)j[0]["InvImp"] + (decimal)r["InvImp"];
                        if ((decimal)j[0]["InvImp"] > 0 && (decimal)j[0]["inv_qta"] > 0)
                            j[0]["inv_cos"] = (decimal)j[0]["InvImp"] / (decimal)j[0]["inv_qta"];
                    }

                }
            }
            else if (strTip == "ERR")
            {
                DataView v = new DataView(((DataView)dgv1.DataSource).ToTable(), "", "inv_nri", DataViewRowState.CurrentRows);
                t = v.ToTable();
            }

            if(t.Rows.Count == 0)
                MessageBox.Show("Non sono presenti dati da estrarre!", "INVENTARIO SU PDF", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            else
            {
                // Create a new PDF document
                PdfDocument pd = new PdfDocument();
                pd.Info.Title = "Articoli";

                // Create an empty page
                PdfPage page = pd.AddPage();

                // Get an XGraphics object for drawing
                XGraphics gfx = XGraphics.FromPdfPage(page);

                //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

                // Create a font
                XFont font1 = new XFont("Courier new", 10, XFontStyle.Bold);
                XFont font2 = new XFont("Courier new", 9, XFontStyle.Regular);
                XFont font3 = new XFont("Courier new", 8, XFontStyle.Regular);
                XFont font4 = new XFont("Courier new", 8, XFontStyle.Bold);

                double Y = 5;
                double X = 20;

                XStringFormat frmDX = new XStringFormat();
                frmDX.Alignment = XStringAlignment.Far;
                frmDX.LineAlignment = XLineAlignment.Far;

                string sFam = ""; // (string)t.Rows[0]["InvL1c"] + (string)t.Rows[0]["InvL2c"] + (string)t.Rows[0]["InvL3c"];

                decimal dRig = 0;
                decimal dPez = 0;
                decimal dVal = 0;
                decimal dTotRig = 0;
                decimal dTotPez = 0;
                decimal dTotVal = 0;

                progressBar1.Value = 0;
                progressBar1.Maximum = t.Rows.Count;
                progressBar1.Minimum = 0;

                for (int i = 0; i <= t.Rows.Count - 1; i++)
                {
                    progressBar1.Increment(1);
                    System.Windows.Forms.Application.DoEvents();

                    if (iRow > iRows || iRow == 0)
                    {
                        if (iRow > 0)
                        {
                            page = pd.AddPage();
                            gfx = XGraphics.FromPdfPage(page);
                        }

                        X = 20;
                        iRow = 0;
                    
                        s = txtIntDes.Text;
                        if(s == "")
                            s = "STAMPA INVENTARIO VALORIZZATO AL " + dtpIntDay.Value.ToString("dd/MM/yyyy");
                        if (rdbPrnQtaGia.Checked)
                            s += " su giacenza";
                        if(sRil != "")
                            s += " Terminale " + sTer + " Rilevazione " + sRil;

                        //if (chkFor.Checked)
                        //    s += "FORNITORI ";
                        gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                        X += 20;

                        s = "Codice";
                        gfx.DrawString(s, font1, XBrushes.Black, Y + 60, X, XStringFormats.Default);

                        s = "Descrizione";
                        gfx.DrawString(s, font1, XBrushes.Black, Y + 100, X, XStringFormats.Default);

                        s = "UM";
                        gfx.DrawString(s, font1, XBrushes.Black, Y + 325, X, XStringFormats.Default);

                        s = "Q.ta";
                        gfx.DrawString(s, font1, XBrushes.Black, Y + 365, X, XStringFormats.Default);

                        s = "Costo";
                        gfx.DrawString(s, font1, XBrushes.Black, Y + 400, X, XStringFormats.Default);

                        s = "Importo";
                        gfx.DrawString(s, font1, XBrushes.Black, Y + 450, X, XStringFormats.Default);

                        X += 15;
                    }

                    iRow++;

                    if (t.Rows.Count - 1 >= i)
                    {
                        if (!DBNull.Value.Equals(t.Rows[i]["InvL1c"]) && !DBNull.Value.Equals(t.Rows[i]["InvL1c"]) && !DBNull.Value.Equals(t.Rows[i]["InvL1c"]))
                        {
                            if (sFam != (string)t.Rows[i]["InvL1c"] + (string)t.Rows[i]["InvL2c"] + (string)t.Rows[i]["InvL3c"])
                            {
                                if (i > 0)
                                {
                                    s = "Famiglia totali righe " + dRig.ToString() + " ";
                                    gfx.DrawString(s, font3, XBrushes.Black, Y + 10, X, XStringFormats.Default);
                                    s = (dPez).ToString("###0.00");
                                    gfx.DrawString(s, font3, XBrushes.Black, Y + 390, X, frmDX);
                                    s = dVal.ToString("###,##0.00");
                                    gfx.DrawString(s, font3, XBrushes.Black, Y + 480, X, frmDX);
                                    X += 20;
                                    iRow++;
                                }

                                dRig = 0;
                                dPez = 0;
                                dVal = 0;

                                sFam = (string)t.Rows[i]["InvL1c"] + (string)t.Rows[i]["InvL2c"] + (string)t.Rows[i]["InvL3c"];

                                s = "Famiglia " + (string)t.Rows[i]["InvL1c"] + " " +
                                    ((string)t.Rows[i]["InvL1d"]).Trim() + " - " +
                                    (string)t.Rows[i]["InvL2c"] + " " +
                                    ((string)t.Rows[i]["InvL2d"]).Trim() + " - " +
                                    (string)t.Rows[i]["InvL3c"] + " " +
                                    ((string)t.Rows[i]["InvL3d"]).Trim();
                                gfx.DrawString(s, font3, XBrushes.Black, Y + 10, X, XStringFormats.Default);
                                X += 20;
                                iRow++;
                            }
                        }
                        s = "";
                        if (!DBNull.Value.Equals(t.Rows[i]["inv_art"]))
                            s = (string)t.Rows[i]["inv_art"];
                        gfx.DrawString(s, font4, XBrushes.Black, Y + 60, X, XStringFormats.Default);

                        s = "";
                        if (!DBNull.Value.Equals(t.Rows[i]["InvArd"]))
                            s = (string)t.Rows[i]["InvArd"];
                        gfx.DrawString(s, font4, XBrushes.Black, Y + 100, X, XStringFormats.Default);

                        s = "";
                        if (!DBNull.Value.Equals(t.Rows[i]["InvUmi"]))
                            s = (string)t.Rows[i]["InvUmi"];
                        gfx.DrawString(s, font4, XBrushes.Black, Y + 325, X, XStringFormats.Default);

                        s = "";

                        if (!DBNull.Value.Equals(t.Rows[i]["inv_qta"]))
                            s = ((decimal)t.Rows[i]["inv_qta"]).ToString("####,##0.00");
                        gfx.DrawString(s, font4, XBrushes.Black, Y + 390, X + 3, frmDX);

                        s = "";
                        if (!DBNull.Value.Equals(t.Rows[i]["inv_cos"]))
                            s = ((decimal)t.Rows[i]["inv_cos"]).ToString("####,##0.000");
                        gfx.DrawString(s, font4, XBrushes.Black, Y + 430, X + 3, frmDX);

                        s = "";
                        if (!DBNull.Value.Equals(t.Rows[i]["InvImp"]))
                            s = ((decimal)t.Rows[i]["InvImp"]).ToString("####,##0.000");
                        gfx.DrawString(s, font4, XBrushes.Black, Y + 480, X + 3, frmDX);

                        dRig++;
                        dPez += (decimal)t.Rows[i]["inv_qta"];
                        dVal += (decimal)t.Rows[i]["InvImp"];
                        dTotRig++;
                        dTotPez += (decimal)t.Rows[i]["inv_qta"];
                        dTotVal += (decimal)t.Rows[i]["InvImp"];
                    }

                    X += 20;
                    //break;

                }

                s = "Famiglia totali righe " + dRig.ToString() + " ";
                gfx.DrawString(s, font3, XBrushes.Black, Y + 10, X, XStringFormats.Default);
                s = (dPez).ToString("###0.00");
                gfx.DrawString(s, font3, XBrushes.Black, Y + 390, X, frmDX);
                s = dVal.ToString("###,##0.00");
                gfx.DrawString(s, font3, XBrushes.Black, Y + 480, X, frmDX);

                X += 25;

                s = "Totale righe " + dTotRig.ToString("###,##0");
                s += "    Totale pezzi " + dTotPez.ToString("###,##0.00");
                s += "    Totale valore costi inventario   " + dTotVal.ToString("###,##0.00");
                gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                // Save the document...
                try
                {
                    string sFil = "C:\\ApProject\\PDF\\Inv_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                    pd.Save(sFil);
                    // ...and start a viewer.
                    Process.Start(sFil);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("File di stampa già aperto!");
                }
            }
        }

        private void btnXls_Click(object sender, EventArgs e)
        {
            Excel1();
        }

        private void btnDiff_Click(object sender, EventArgs e)
        {
            ExcelDiff();
        }

        private void oldExcel1()
        {
            string s = "";
            decimal d = 0;
            DataRow[] j;
            DataTable t = new DataTable();
            DataView v = new DataView(((DataView)dgv1.DataSource).ToTable(), "inv_msg ='' AND inv_art<>''", "InvL1d,InvL2d,InvL3d,inv_art", DataViewRowState.CurrentRows);
            t = v.Table.Clone();

            DataTable tIve = new clsGenTabTmp().TabTmpIvaRiep("TabIva");

            s = "SELECT * FROM TabIva";
            DataTable tIva = _clsFun.FillTabSql("TabIva", s, false, _strConSql);

            progressBar1.Value = 0;
            progressBar1.Maximum = v.Count;
            progressBar1.Minimum = 0;

            foreach (DataRowView r in v)
            {
                progressBar1.Increment(1);
                System.Windows.Forms.Application.DoEvents();


                if (DBNull.Value.Equals(r["InvIva"]))
                    r["InvIva"] = "";

                if ((string)r["inv_art"] == "0003984")
                    Console.WriteLine("aaaa");

                j = t.Select("inv_art='" + r["inv_art"] + "'");
                if (j.Length == 0)
                {
                    DataRow x = t.NewRow();
                    x["inv_art"] = r["inv_art"];
                    x["InvArd"] = r["InvArd"];
                    x["InvIva"] = r["InvIva"];
                    x["InvL1c"] = r["InvL1c"];
                    x["InvL2c"] = r["InvL2c"];
                    x["InvL3c"] = r["InvL3c"];
                    x["InvL1d"] = r["InvL1d"];
                    x["InvL2d"] = r["InvL2d"];
                    x["InvL3d"] = r["InvL3d"];
                    x["InvUmi"] = r["InvUmi"];
                    x["inv_qta"] = 0;
                    x["inv_cos"] = 0;
                    x["InvImp"] = 0;
                    t.Rows.Add(x);
                    j = t.Select("inv_art='" + r["inv_art"] + "'");
                }

                j[0]["inv_qta"] = (decimal)j[0]["inv_qta"] + (decimal)r["inv_qta"];
                j[0]["InvImp"] = (decimal)j[0]["InvImp"] + (decimal)r["InvImp"];
                if ((decimal)j[0]["InvImp"] > 0 && (decimal)j[0]["inv_qta"] > 0)
                    j[0]["inv_cos"] = (decimal)j[0]["InvImp"] / (decimal)j[0]["inv_qta"];

                d += (decimal)r["InvImp"];

                j = tIve.Select("iva_cod='" + (string)r["InvIva"] + "'");
                if(j.Length == 0)
                {
                    DataRow x = tIve.NewRow();
                    x["iva_cod"] = r["InvIva"];

                    j = tIva.Select("tab_cod='" + r["InvIva"] + "'");
                    if(j.Length > 0 )
                    {
                        x["iva_des"] = (string)j[0]["tab_des"];
                        x["iva_ali"] = Convert.ToDecimal(j[0]["tab_ali"]);
                    }
                    x["iva_iva"] = 0;
                    x["iva_imp"] = 0;
                    x["iva_tot"] = 0;
                    tIve.Rows.Add(x);
                }

                j = tIve.Select("iva_cod='" + r["InvIva"] + "'");

                d = _clsFun.ValIva((decimal)r["InvImp"], Convert.ToDecimal(j[0]["iva_ali"]));

                j[0]["iva_iva"] = (decimal)j[0]["iva_iva"] + d;
                j[0]["iva_imp"] = (decimal)j[0]["iva_imp"] + (decimal)r["InvImp"];
                j[0]["iva_tot"] = (decimal)j[0]["iva_tot"] + (decimal)r["InvImp"] + d;

            }

            foreach(DataRow y in tIve.Rows)
            {
                DataRow x = t.NewRow();
                x["inv_art"] = "";
                x["InvArd"] = y["iva_des"];
                x["InvIva"] = y["iva_cod"];
                x["InvL1c"] = "";
                x["InvL2c"] = "";
                x["InvL3c"] = "";
                x["InvL1d"] = "";
                x["InvL2d"] = "";
                x["InvL3d"] = "";
                x["InvUmi"] = "";
                x["inv_qta"] = y["iva_iva"];
                x["inv_cos"] = y["iva_imp"];
                x["InvImp"] = y["iva_tot"];
                t.Rows.Add(x);

                d += (decimal)y["iva_tot"];
            }

            string sTit = "INVENTARIO AL " +  dtpIntDay.Value.ToShortDateString() +  " - " + txtIntDes.Text;
            string sFil = "";

            s = "";
            s += "InvL1c, Liv 1, 30, StringLiteral;";
            s += "InvL1d, Liv 1 des, 70, StringLiteral;";
            s += "InvL2c, Liv 2, 30, StringLiteral;";
            s += "InvL2d, Liv 2 des, 90, StringLiteral;";
            s += "InvL3c, Liv 3, 30, StringLiteral;";
            s += "InvL3d, Liv 3 des, 90, StringLiteral;";
            s += "inv_art, Articolo, 50, StringLiteral;";
            s += "InvArd, Descrizione, 180, StringLiteral;";
            s += "InvUmi, UM, 20, StringLiteral;";
            s += "InvIva, IVA, 20, Decimal2;";
            s += "inv_qta, Q.tà, 40, Decimal2;";
            s += "inv_cos, Prezzo, 40, Decimal3;";
            s += "InvImp, Importo, 50, Decimal3;";

            for (int i = 0; i <= 100; i++)
            {
                sFil = "C:\\APproject\\XLS\\" + DateTime.Now.ToString("yyyyMMdd") + "_Inventario.xls";
                if (File.Exists(sFil))
                {
                    try
                    {
                        File.Delete(sFil);
                        break;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }

                }
                else
                    break;
            }

            DataSet ds = new DataSet();

            DataTable tTit = (new clsExcel()).FillTitStr(s, "TabTit");

            ds.Tables.Add(tTit);
            ds.Tables.Add(t);

            (new clsExcel()).exportToExcel(ds, sFil, sTit, "Totale " + d.ToString("#,###,##0.00"), true);

            ds.Tables.Remove(t);
            ds.Clear();
            ds.Dispose();
        }

        private void Excel1()
        {
            string s = "";
            decimal d = 0;
            DataRow[] j;
            DataTable t = new DataTable();
            DataView v = new DataView(((DataView)dgv1.DataSource).ToTable(), "inv_msg ='' AND inv_art<>''", "InvL1d,InvL2d,InvL3d,inv_art", DataViewRowState.CurrentRows);
            t = v.Table.Clone();

            DataTable tTmp = new clsGenTabTmp().TabTmpArt("TmpArt");

            DataTable tIve = new clsGenTabTmp().TabTmpIvaRiep("TabIva");

            s = "SELECT * FROM TabIva";
            DataTable tIva = _clsFun.FillTabSql("TabIva", s, false, _strConSql);

            progressBar1.Value = 0;
            progressBar1.Maximum = v.Count;
            progressBar1.Minimum = 0;

            foreach (DataRowView r in v)
            {
                progressBar1.Increment(1);
                System.Windows.Forms.Application.DoEvents();

                /*** Prezzo di vendita ***/
                tTmp.Rows.Clear();
                DataRow x = tTmp.NewRow();
                x["tmp_art"] = r["inv_art"];
                tTmp.Rows.Add(x);
                tTmp = _clsQry.ArtPrezzo(tTmp, "");
                r["InvPrv"] = tTmp.Rows[0]["tmp_prv"];


                Boolean b = true;

                if (rdbPrnQtaGia.Checked && (decimal)r["inv_gia"] == 0)
                    b = false;
                else if (rdbPrnQtaRil.Checked && (decimal)r["inv_qta"] == 0)
                    b = false;

                if (b)
                {
                    if (DBNull.Value.Equals(r["InvIva"]))
                        r["InvIva"] = "";

                    if ((string)r["inv_art"] == "0003984")
                        Console.WriteLine("aaaa");

                    j = t.Select("inv_art='" + r["inv_art"] + "'");
                    if (j.Length == 0)
                    {
                        x = t.NewRow();
                        x["inv_art"] = r["inv_art"];
                        x["InvArd"] = r["InvArd"];
                        x["InvIva"] = r["InvIva"];
                        x["InvL1c"] = r["InvL1c"];
                        x["InvL2c"] = r["InvL2c"];
                        x["InvL3c"] = r["InvL3c"];
                        x["InvL1d"] = r["InvL1d"];
                        x["InvL2d"] = r["InvL2d"];
                        x["InvL3d"] = r["InvL3d"];
                        x["InvUmi"] = r["InvUmi"];
                        x["inv_qta"] = 0;
                        x["inv_gia"] = 0;
                        x["inv_cos"] = 0; 
                        x["InvImp"] = 0;
                        x["InvPrv"] = r["InvPrv"];
                        t.Rows.Add(x);
                        j = t.Select("inv_art='" + r["inv_art"] + "'");
                    }

                    if (rdbPrnQtaGia.Checked)
                    {
                        j[0]["inv_gia"] = (decimal)j[0]["inv_gia"] + (decimal)r["inv_gia"];
                        j[0]["InvImp"] = (decimal)j[0]["InvImp"] + ((decimal)r["inv_gia"] * (decimal)r["inv_cos"]);
                        if ((decimal)j[0]["InvImp"] > 0 && (decimal)j[0]["inv_gia"] > 0)
                            j[0]["inv_cos"] = (decimal)j[0]["InvImp"] / (decimal)j[0]["inv_gia"];
                        r["InvImp"] = j[0]["InvImp"];
                    }
                    else
                    {
                        j[0]["inv_qta"] = (decimal)j[0]["inv_qta"] + (decimal)r["inv_qta"];
                        j[0]["InvImp"] = (decimal)j[0]["InvImp"] + (decimal)r["InvImp"];
                        if ((decimal)j[0]["InvImp"] > 0 && (decimal)j[0]["inv_qta"] > 0)
                            j[0]["inv_cos"] = (decimal)j[0]["InvImp"] / (decimal)j[0]["inv_qta"];
                    }

                    d += (decimal)r["InvImp"];

                    j = tIve.Select("iva_cod='" + (string)r["InvIva"] + "'");
                    if (j.Length == 0)
                    {
                        x = tIve.NewRow();
                        x["iva_cod"] = r["InvIva"];

                        j = tIva.Select("tab_cod='" + r["InvIva"] + "'");
                        if (j.Length > 0)
                        {
                            x["iva_des"] = (string)j[0]["tab_des"];
                            x["iva_ali"] = Convert.ToDecimal(j[0]["tab_ali"]);
                        }
                        x["iva_iva"] = 0;
                        x["iva_imp"] = 0;
                        x["iva_tot"] = 0;
                        tIve.Rows.Add(x);
                    }

                    j = tIve.Select("iva_cod='" + r["InvIva"] + "'");

                    d = _clsFun.ValIva((decimal)r["InvImp"], Convert.ToDecimal(j[0]["iva_ali"]));

                    j[0]["iva_iva"] = (decimal)j[0]["iva_iva"] + d;
                    j[0]["iva_imp"] = (decimal)j[0]["iva_imp"] + (decimal)r["InvImp"];
                    j[0]["iva_tot"] = (decimal)j[0]["iva_tot"] + (decimal)r["InvImp"] + d;
                }
            }

            foreach (DataRow y in tIve.Rows)
            {
                DataRow x = t.NewRow();
                x["inv_art"] = "";
                x["InvArd"] = y["iva_des"];
                x["InvIva"] = y["iva_cod"];
                x["InvL1c"] = "";
                x["InvL2c"] = "";
                x["InvL3c"] = "";
                x["InvL1d"] = "";
                x["InvL2d"] = "";
                x["InvL3d"] = "";
                x["InvUmi"] = "";
                x["inv_qta"] = y["iva_iva"];
                x["inv_cos"] = y["iva_imp"];
                x["InvImp"] = y["iva_tot"];
                t.Rows.Add(x);

                d += (decimal)y["iva_tot"];
            }

            if(t.Rows.Count == 0)
                MessageBox.Show("Non sono presenti dati da estrarre!", "INVENTARIO SU EXCEL", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            else
            {
                string sTit = "INVENTARIO AL " + dtpIntDay.Value.ToShortDateString() + " - " + txtIntDes.Text;
                if (rdbPrnQtaGia.Checked)
                    sTit += " su giacenza";
                string sFil = "";
                string sFoo = ",,,,,,,'Totali',,,,," + d.ToString("#0.00").Replace(",", ".") + ",";

                string sFld = "";
                sFld += "InvL1c, Liv 1, 30, StringLiteral;";
                sFld += "InvL1d, Liv 1 des, 70, StringLiteral;";
                sFld += "InvL2c, Liv 2, 30, StringLiteral;";
                sFld += "InvL2d, Liv 2 des, 90, StringLiteral;";
                sFld += "InvL3c, Liv 3, 30, StringLiteral;";
                sFld += "InvL3d, Liv 3 des, 90, StringLiteral;";
                sFld += "inv_art, Articolo, 50, StringLiteral;";
                sFld += "InvArd, Descrizione, 180, StringLiteral;";
                sFld += "InvUmi, UM, 20, StringLiteral;";
                sFld += "InvIva, IVA, 20, StringLiteral;";

                if (rdbPrnQtaGia.Checked)
                    sFld += "inv_gia, Q.tà, 40, Decimal;";
                else
                    sFld += "inv_qta, Q.tà, 40, Decimal;";

                sFld += "inv_cos, Costo, 40, Decimal;";
                sFld += "InvImp, Importo, 50, Decimal;";
                sFld += "InvPrv, Prezzo vendita, 40, Decimal;";

                sFil = "Inventario.xls";

                (new clsExcel2()).exportToXls1(t, sFld, sFil, sTit, sFoo);
            }
        }

        private void ExcelDiff()
        {
            string s = "";
            decimal d = 0;
            DataRow[] j;
            DataTable t = new DataTable();
            DataView v = new DataView(((DataView)dgv1.DataSource).ToTable(), "inv_msg ='' AND inv_art<>''", "InvL1d,InvL2d,InvL3d,inv_art", DataViewRowState.CurrentRows);
            t = v.Table.Clone();

            DataTable tIve = new clsGenTabTmp().TabTmpIvaRiep("TabIva");

            s = "SELECT * FROM TabIva";
            DataTable tIva = _clsFun.FillTabSql("TabIva", s, false, _strConSql);

            progressBar1.Value = 0;
            progressBar1.Maximum = v.Count;
            progressBar1.Minimum = 0;

            foreach (DataRowView r in v)
            {
                progressBar1.Increment(1);
                System.Windows.Forms.Application.DoEvents();

                if ((decimal)r["InvDif"] != 0)
                {

                    if (DBNull.Value.Equals(r["InvIva"]))
                        r["InvIva"] = "";

                    if ((string)r["inv_art"] == "0010779")
                        Console.WriteLine("aaaa");

                    j = t.Select("inv_art='" + r["inv_art"] + "'");
                    if (j.Length == 0)
                    {
                        DataRow x = t.NewRow();
                        x["inv_art"] = r["inv_art"];
                        x["InvArd"] = r["InvArd"];
                        x["InvIva"] = r["InvIva"];
                        x["InvL1c"] = r["InvL1c"];
                        x["InvL2c"] = r["InvL2c"];
                        x["InvL3c"] = r["InvL3c"];
                        x["InvL1d"] = r["InvL1d"];
                        x["InvL2d"] = r["InvL2d"];
                        x["InvL3d"] = r["InvL3d"];
                        x["InvUmi"] = r["InvUmi"];
                        x["inv_qta"] = 0;
                        x["inv_gia"] = 0;
                        x["InvDif"] = 0;
                        t.Rows.Add(x);
                        j = t.Select("inv_art='" + r["inv_art"] + "'");
                    }

                    j[0]["inv_gia"] = (decimal)j[0]["inv_gia"] + (decimal)r["inv_gia"];
                    j[0]["inv_qta"] = (decimal)j[0]["inv_qta"] + (decimal)r["inv_qta"];
                    j[0]["InvDif"] = (decimal)j[0]["inv_qta"] - (decimal)j[0]["inv_gia"];
                    if ((decimal)j[0]["InvImp"] > 0 && (decimal)j[0]["inv_qta"] > 0)
                        j[0]["inv_cos"] = (decimal)j[0]["InvImp"] / (decimal)j[0]["inv_qta"];

                    d += (decimal)r["InvImp"];

                    j = tIve.Select("iva_cod='" + (string)r["InvIva"] + "'");
                    if (j.Length == 0)
                    {
                        DataRow x = tIve.NewRow();
                        x["iva_cod"] = r["InvIva"];

                        j = tIva.Select("tab_cod='" + r["InvIva"] + "'");
                        if (j.Length > 0)
                        {
                            x["iva_des"] = (string)j[0]["tab_des"];
                            x["iva_ali"] = Convert.ToDecimal(j[0]["tab_ali"]);
                        }
                        x["iva_iva"] = 0;
                        x["iva_imp"] = 0;
                        x["iva_tot"] = 0;
                        tIve.Rows.Add(x);
                    }

                    j = tIve.Select("iva_cod='" + r["InvIva"] + "'");

                    d = _clsFun.ValIva((decimal)r["InvImp"], Convert.ToDecimal(j[0]["iva_ali"]));

                    j[0]["iva_iva"] = (decimal)j[0]["iva_iva"] + d;
                    j[0]["iva_imp"] = (decimal)j[0]["iva_imp"] + (decimal)r["InvImp"];
                    j[0]["iva_tot"] = (decimal)j[0]["iva_tot"] + (decimal)r["InvImp"] + d;
                }
            }

            foreach (DataRow y in tIve.Rows)
            {
                DataRow x = t.NewRow();
                x["inv_art"] = "";
                x["InvArd"] = y["iva_des"];
                x["InvIva"] = y["iva_cod"];
                x["InvL1c"] = "";
                x["InvL2c"] = "";
                x["InvL3c"] = "";
                x["InvL1d"] = "";
                x["InvL2d"] = "";
                x["InvL3d"] = "";
                x["InvUmi"] = "";
                x["inv_qta"] = y["iva_iva"];
                x["inv_cos"] = y["iva_imp"];
                x["InvImp"] = y["iva_tot"];
                t.Rows.Add(x);

                d += (decimal)y["iva_tot"];
            }

            string sTit = "DIFFERENZE INVENTARIO AL " + dtpIntDay.Value.ToShortDateString() + " - " + txtIntDes.Text;
            string sFil = "";
            string sFoo = ",,,,,,,'Totali',,,,," + d.ToString("#0.00").Replace(",", ".");

            string sFld = "";
            sFld += "InvL1c, Liv 1, 30, StringLiteral;";
            sFld += "InvL1d, Liv 1 des, 70, StringLiteral;";
            sFld += "InvL2c, Liv 2, 30, StringLiteral;";
            sFld += "InvL2d, Liv 2 des, 90, StringLiteral;";
            sFld += "InvL3c, Liv 3, 30, StringLiteral;";
            sFld += "InvL3d, Liv 3 des, 90, StringLiteral;";
            sFld += "inv_art, Articolo, 50, StringLiteral;";
            sFld += "InvArd, Descrizione, 180, StringLiteral;";
            sFld += "InvUmi, UM, 20, StringLiteral;";
            sFld += "InvIva, IVA, 20, StringLiteral;";
            sFld += "inv_gia, Giacenza, 40, Decimal;";
            sFld += "inv_qta, Q.tà, 40, Decimal;";
            sFld += "InvDif, Differenze, 40, Decimal;";
            //sFld += "inv_cos, Prezzo, 40, Decimal;";
            //sFld += "InvImp, Importo, 50, Decimal;";

            sFil = "InvDiff.xls";

            (new clsExcel2()).exportToXls1(t, sFld, sFil, sTit, sFoo);
        }

        private void btnAggCosti_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Confermi il controllo per articoli con costo a ZERO?", "AGGIORNAMENTO COSTI", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes) 
                AggCosti();
        }

        private void AggCosti()
        {
            DataTable t = ((DataView)dgv1.DataSource).Table;

            DataTable tArt = new DataTable("TabArt");

            foreach(DataRow y in t.Rows)
            {
                if((decimal)y["inv_cos"] == 0)
                {
                    tArt = _clsQry.ArtSeek((string)y["inv_art"], "SEEK");
                    if (tArt.Rows.Count > 0 && !DBNull.Value.Equals(tArt.Rows[0]["tmp_cos"]) && (decimal)tArt.Rows[0]["tmp_cos"] > 0)
                    {
                        y["inv_cos"] = (decimal)tArt.Rows[0]["tmp_cos"]; 
                        y["InvMdy"] = "S";
                    }
                }
            }
        }

        private void simulatoreToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmGesInvSimula f = new frmGesInvSimula();
            f.ShowDialog();

            if (f._decVal > 0 && f._decArt > 0)
                FillInvSimulato(f._decVal, f._decArt);
        }

        private void FillInvSimulato(decimal decVal, decimal decArt)
        {
            DataTable tEcr = _clsQry.tabEcr();
            string s = "SELECT * FROM TabReparti";
            DataTable tRep = _clsFun.FillTabSql("TabRep", s, false, _strConSql);
            DataTable tInv = ((DataView)dgv1.DataSource).Table;
            DataTable t = new DataTable();
            DataRow[] j;

            string sNte = "01";
            string sNrv = "01";
            int i = tInv.Rows.Count;
            int iArt = 0;
            int iQta = 0;
            Decimal dVal = 0;

            Random rnd = new Random();

            s = "SELECT * FROM AnaArticoli WHERE art_sta='A' ORDER BY art_dtm";
            if(cmbIntRep.SelectedValue != null && cmbIntRep.SelectedValue.ToString() != "")
                s = "SELECT * FROM AnaArticoli WHERE art_sta='A' AND art_rep='" + cmbIntRep.SelectedValue.ToString() + "' ORDER BY art_dtm";

            DataTable tArt = _clsFun.FillTabSql("AnaArticoli", s, false, _strConSql);

            progressBar1.Value = 0;
            progressBar1.Maximum = Convert.ToInt32(decVal);
            progressBar1.Minimum = 0;


            foreach (DataRow y in tArt.Rows)
            {
                j = tInv.Select("inv_art='" + y["art_cod"] + "'");
                if(j.Length == 0)
                {
                    t = _clsQry.ArtSeek((string)y["art_cod"], "");
                    if(t.Rows.Count > 0)
                    {
                        DataRow x = tInv.NewRow();
                        x["inv_num"] = lblIntNum.Text;
                        x["inv_nte"] = sNte;
                        x["inv_nrv"] = sNrv;
                        x["inv_nri"] = i.ToString("00000");
                        x["inv_ean"] = "";      // (string)y["ord_ean"];
                        x["inv_gia"] = 0;
                        x["inv_qta"] = 0;       // (decimal)y["ord_qta"];
                        x["InvDif"] = 0;
                        x["inv_art"] = (string)t.Rows[0]["tmp_art"];
                        x["InvArd"] = (string)t.Rows[0]["tmp_ard"];
                        x["InvUmi"] = (string)t.Rows[0]["tmp_umi"];
                        x["inv_cos"] = (decimal)t.Rows[0]["tmp_cos"];
                        x["inv_msg"] = "";

                        s = "l1c+l2c+l3c='" + (string)t.Rows[0]["tmp_ecr"] + "'";
                        j = tEcr.Select(s);
                        if (j.Length > 0)
                        {
                            x["InvL1d"] = (string)j[0]["l1d"];
                            x["InvL2d"] = (string)j[0]["l2d"];
                            x["InvL3d"] = (string)j[0]["l3d"];
                        }

                        s = "tab_cod='" + (string)t.Rows[0]["tmp_rep"] + "'";
                        j = tRep.Select(s);
                        if (j.Length > 0)
                            x["InvRed"] = (string)j[0]["tab_des"];
                        x["InvMdy"] = "S";
                        tInv.Rows.Add(x);

                        i++;
                    }
                }

                iQta = rnd.Next(1, 17); 

                j = tInv.Select("inv_art='" + y["art_cod"] + "'");
                j[0]["inv_qta"] = (decimal)j[0]["inv_qta"] + iQta;
                j[0]["InvImp"] = (decimal)j[0]["inv_qta"] * (decimal)j[0]["inv_cos"];

                dVal += (decimal)j[0]["inv_cos"] * iQta;

                if (dVal < progressBar1.Maximum)
                {
                    progressBar1.Value = Convert.ToInt32(dVal);
                    System.Windows.Forms.Application.DoEvents();
                }

                if (iArt > decArt || dVal > decVal)
                    break;
            }

            if (dVal < decVal)
            {
                foreach (DataRow y in tInv.Rows)
                {
                    iQta = rnd.Next(1, 17); 

                    y["inv_qta"] = (decimal)y["inv_qta"] + iQta;
                    y["InvImp"] = (decimal)y["inv_qta"] * (decimal)y["inv_cos"];

                    dVal += (decimal)y["inv_cos"] * iQta;

                    //progressBar1.Value = Convert.ToInt32(dVal); //.Increment(Convert.ToInt32((decimal)j[0]["inv_cos"] * iQta));

                    if(dVal < progressBar1.Maximum)
                        progressBar1.Value = Convert.ToInt32(dVal); //.Increment(Convert.ToInt32((decimal)j[0]["inv_cos"] * iQta));

                    System.Windows.Forms.Application.DoEvents();


                    if (iArt > decArt || dVal > decVal)
                        break;
                }
            }

            Totali();
            NewRiga(true);
        }

        private void txtSeek_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                string s = txtSeek.Text.Trim();

                string sFld = "Descrizione";
                if (_clsFun.Numerico(s) && s.Length < 8)
                    sFld = "Articolo";
                else if (_clsFun.Numerico(s))
                    sFld = "Barcode";

                if (s != "" && dgv1.RowCount > 0)
                {
                    Boolean b = false;
                    int iCur = dgv1.CurrentCell.RowIndex;
                    int i = 0;

                    foreach (DataGridViewRow row in dgv1.Rows)
                    {
                        i++;

                        if (i > iCur + 1)
                        {
                            //if (row.Cells["Cliente/Fornitore"].Value.ToString().ToLower().Contains(s.ToLower()))
                            if (row.Cells[sFld].Value.ToString().ToLower().Contains(s.ToLower()))
                            {
                                dgv1.Rows[row.Index].Selected = true;
                                dgv1.FirstDisplayedScrollingRowIndex = row.Index;
                                dgv1.CurrentCell = dgv1.Rows[row.Index].Cells[0];
                                b = true;
                                break;
                            }
                        }
                    }
                    if (!b)
                        dgv1.CurrentCell = dgv1.Rows[0].Cells[0];
                }
            }
        }

        private void importAPPhoneToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string s = "SELECT * FROM TabForImport WHERE tab_tip='FTPDIV'";
            DataTable t = _clsFun.FillTabSql("TabForImp", s, true, _strConSql);
            if (t.Rows.Count > 0)
                new frmGesForDivulgazioni().ForDownLoad();
            else
                MessageBox.Show("Opzione non attiva!", "IMPORTAZIONE DA APSHOP", MessageBoxButtons.OK, MessageBoxIcon.Question);
        }

        private void btnGia_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Caricamento giacenza, confermi?", this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
            {
                FillGiacenza();
            }
        }

        private void FillGiacenza()
        {
            string sChoRep = cmbIntRep.SelectedValue.ToString();

            DataRow[] j;
            DataTable tEcr = _clsQry.tabEcr();
            DataTable tRep = _clsFun.FillTabSql("TabRep", "SELECT * FROM TabReparti", false, _strConSql);

            DataTable tInv = ((DataView)dgv1.DataSource).Table;

            int i = tInv.Rows.Count;

            string s = "SELECT * FROM AnaArtGiacenza";
            DataTable t = _clsFun.FillTabSql("AnaArtGiacenza", s, false, _strConSql);

            progressBar1.Value = 0;
            progressBar1.Maximum = (t.Rows.Count);
            progressBar1.Minimum = 0;

            foreach(DataRow y in t.Rows)
            {
                progressBar1.Increment(1);
                System.Windows.Forms.Application.DoEvents();

                string sArt = (string)y["gia_art"];
                string sNte = "00";
                string sNrv = "00";
                //decimal dGia = (decimal)y["gia_gia"];

                string sRed = "";

                DataTable tArt = _clsQry.ArtSeek(sArt, "");
                if (tArt.Rows.Count > 0 && (sChoRep == "" || sChoRep == (string)tArt.Rows[0]["tmp_rep"]))
                { 
                    s = "tab_cod='" + (string)tArt.Rows[0]["tmp_rep"] + "'";
                    j = tRep.Select(s);
                    if(j.Length > 0)
                        sRed = (string)j[0]["tab_des"];

                    j = tInv.Select("inv_art='" + sArt + "'");
                    if (j.Length == 0)
                    {
                        i++;

                        DataRow x = tInv.NewRow();
                        x["inv_num"] = lblIntNum.Text;
                        x["inv_nte"] = sNte;
                        x["inv_nrv"] = sNrv;
                        x["inv_nri"] = i.ToString("00000");
                        x["inv_ean"] = "";
                        x["inv_qta"] = 0;
                        x["inv_msg"] = "";
                        x["inv_dti"] = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                        x["inv_dtt"] = "";

                        x["inv_art"] = "";
                        x["InvArd"] = "";
                        x["InvL1c"] = "";
                        x["InvL2c"] = "";
                        x["InvL3c"] = "";
                        x["InvUmi"] = "";
                        x["inv_cos"] = 0;

                        x["inv_art"] = sArt;
                        x["InvArd"] = (string)tArt.Rows[0]["tmp_ard"];
                        x["InvUmi"] = (string)tArt.Rows[0]["tmp_umi"];
                        x["inv_cos"] = (decimal)tArt.Rows[0]["tmp_cos"];
                        x["inv_msg"] = "";

                        s = "l1c+l2c+l3c='" + (string)tArt.Rows[0]["tmp_ecr"] + "'";
                        j = tEcr.Select(s);
                        if (j.Length > 0)
                        {
                            x["InvL1d"] = (string)j[0]["l1d"];
                            x["InvL2d"] = (string)j[0]["l2d"];
                            x["InvL3d"] = (string)j[0]["l3d"];
                        }

                        //if (j.Length > 0)
                        x["InvRed"] = sRed;

                        //}
                        //else
                        //{
                        //    x["inv_msg"] = ((string)x["inv_msg"]).Trim() + " Barcode non trovato " + (string)y["ord_ean"];
                        //    if (((string)x["inv_msg"]).Length > 100)
                        //        x["inv_msg"] = ((string)x["inv_msg"]).Substring(0, 100);
                        //}

                        x["InvMdy"] = "S";
                        tInv.Rows.Add(x);

                        j = tInv.Select("inv_art='" + sArt + "'");
                    }

                    j[0]["inv_gia"] = y["gia_gia"];
                }

            }

            NewRiga(true);
        }

        private void rdbPrnQtaRil_CheckedChanged(object sender, EventArgs e)
        {
            Totali();
        }

    }
}
