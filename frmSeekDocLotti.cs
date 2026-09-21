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
    public partial class frmSeekDocLotti : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        public string _strSelect = "N";
        public string _strYea = "";
        public string _strRes = "";
        public string _strArt = "";
        public string _strTip = "";
        //public string _strDoc = "";

        private const string TABDOCLOT = "GesDocLotti";

        private string _strConSql = "";

        clsResize _form_resize;

        private string _strFrmPar = "";

        public frmSeekDocLotti()
        {
            InitializeComponent();

            _form_resize = new clsResize(this);
            this.Load += _Load;
            this.Resize += _Resize;

            new clsGesGraph().SetGraph(this, 0);
        }
        private void _Load(object sender, EventArgs e)
        {
            _form_resize._get_initial_size();
        }
        private void _Resize(object sender, EventArgs e)
        {
            _form_resize._resize();
        }
        private void frmGesDocLotti_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");

            Video();
            SetDgv1();
            FillTabs();
            FillDati();
        }
        private void Video()
        {
            _form_resize._get_initial_size();
            _strFrmPar = _clsQry.ParForm(this, "R", "");

            this.CenterToScreen();
        }
        private void frmGesDocLotti_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Escape)
                Esci();
        }        
        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }
        private void Esci()
        {
            _clsQry.ParForm(this, "W", "");

            Boolean b = Salva();
            if (!b)
            {
                if (MessageBox.Show("Continui l'uscita?", this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    this.Close();
            }
            else this.Close();
        }

        private void SetDgv1()
        {
            string s = "";

            s = "SELECT * FROM TabStato ORDER BY tab_des";
            DataTable tSta = _clsFun.FillTabSql("TabStato", s, false, _strConSql);

            //s = "SELECT * FROM TabUmi ORDER BY tab_des";
            //DataTable tUmi = _clsFun.FillTabSql(TABTABUMI, s, false, _strConSql);

            dgv1.AutoGenerateColumns = false;
            //dgv1.VirtualMode = true;
            //dgv1.Dock = DockStyle.Fill;
            dgv1.AllowUserToAddRows = false;
            dgv1.ReadOnly = false;
            dgv1.AllowUserToDeleteRows = false;
            //dgv1.DisplayedRowCount() = true;

            DataGridViewTextBoxColumn cTbc;
            //DataGridViewCheckBoxColumn cCbc;
            DataGridViewComboBoxColumn cCmb;

            cCmb = new DataGridViewComboBoxColumn();
            cCmb.DataPropertyName = "lot_sta";
            cCmb.Name = "Stato";
            cCmb.Width = 30;
            cCmb.DataSource = tSta;
            cCmb.ValueMember = "tab_cod";
            cCmb.DisplayMember = "tab_des";
            cCmb.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing;
            dgv1.Columns.Add((DataGridViewColumn)cCmb);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "lot_tip";
            cTbc.Name = "Tipo";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.MaxInputLength = 15;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "lot_cod";
            cTbc.Name = "Lotto";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "ForDes";
            cTbc.Name = "Fornitore";
            cTbc.Width = 100;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.MaxInputLength = 15;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "lot_ndo";
            cTbc.Name = "N.documento";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.MaxInputLength = 15;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "lot_ddo";
            cTbc.Name = "Data";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.MaxInputLength = 15;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "lot_des";
            cTbc.Name = "Descrizione";
            cTbc.Width = 200;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.MaxInputLength = 15;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            //cTbc = new DataGridViewTextBoxColumn();
            //cTbc.DataPropertyName = "lot_idx";
            //cTbc.Name = "iiii";
            //cTbc.Width = 40;
            //cTbc.ValueType = typeof(string);
            //cTbc.ReadOnly = false;
            //cTbc.MaxInputLength = 15;
            ////cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            //dgv1.Columns.Add(cTbc);
        }

        private void FillTabs()
        {
            for (int i = DateTime.Today.Year - 4; i < DateTime.Today.Year + 1; i++)
                cmbYea.Items.Add(i.ToString());
            cmbYea.SelectedIndex = 4;

            string p = "TabStato";
            string s = "SELECT * FROM TabStato WHERE tab_lot=1 ORDER BY tab_cod";
            DataTable t = _clsFun.FillTabSql(p, s, false, _strConSql);
            DataRow x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  Non definito";
            t.Rows.InsertAt(x, 0);
            cmbSta.DataSource = t;
            cmbSta.DisplayMember = "tab_des";
            cmbSta.ValueMember = "tab_cod";
            cmbSta.SelectedValue = "";
        }

        private void FillDati()
        {
            string s = "";

            s = "SELECT ";
            s += "GesDocLotti.*, ";
            //s += "GesDocLotti.lot_idx, ";
            //s += "GesDocLotti.lot_yea, ";
            //s += "GesDocLotti.lot_for, ";
            //s += "GesDocLotti.lot_ndo, ";
            //s += "GesDocLotti.lot_ddo, ";
            //s += "GesDocLotti.lot_cod, ";
            //s += "GesDocLotti.lot_des, ";
            //s += "GesDocLotti.lot_tip, ";
            //s += "GesDocLotti.lot_ann, ";
            //s += "GesDocLotti.lot_nat, ";
            //s += "GesDocLotti.lot_all, ";
            //s += "GesDocLotti.lot_mac, ";
            //s += "GesDocLotti.lot_sez, ";
            //s += "GesDocLotti.lot_lot, ";
            //s += "GesDocLotti.lot_mna, ";
            //s += "GesDocLotti.lot_mdt, ";
            //s += "GesDocLotti.lot_sbo, ";
            //s += "GesDocLotti.lot_sta, ";
            s += "AnaFornitori.for_des AS ForDes ";
            s += "FROM GesDocLotti LEFT OUTER JOIN AnaFornitori ON GesDocLotti.lot_for = AnaFornitori.for_cod ";
            s += "WHERE lot_yea = '" + cmbYea.Text + "' ";
            if (cmbSta.SelectedValue.ToString() != "")
                s += " AND lot_sta='" + cmbSta.SelectedValue.ToString() + "' ";
            s += "ORDER BY CAST(GesDocLotti.lot_cod AS int)";

            DataTable t = _clsFun.FillTabSql(TABDOCLOT, s, false, _strConSql);
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "LotMdy",
                Caption = "Mdy",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });

            dgv1.DataSource = t;

            //dgv1.CurrentCell = dgv1.Rows[dgv1.RowCount - 1].Cells[0];
            if(dgv1.RowCount > 0)
                dgv1.CurrentCell = dgv1.Rows[dgv1.RowCount - 1].Cells[0];
        }

        private void btnNew_Click(object sender, EventArgs e)
        {
            DataTable t = (DataTable)dgv1.DataSource;

            DataRow x = t.NewRow();
            x["lot_cod"] = _clsFun.NewNum(cmbYea.Text, clsDefine.enuNumeratori.NumGesDocLotti, 5, _strConSql);
            x["lot_for"] = "";
            x["lot_ndo"] = "";
            x["lot_ddo"] = DateTime.Today;
            x["lot_tip"] = "";
            x["lot_des"] = "";
            x["lot_nat"] = "";
            x["lot_all"] = "";
            x["lot_mac"] = "";
            x["lot_sez"] = "";
            x["lot_lot"] = "";
            x["lot_mna"] = "";
            x["lot_mdt"] = "";
            x["lot_sbo"] = "";
            x["lot_sta"] = "";
            x["LotMdy"] = "I";

            x["lot_ann"] = false;
            x["lot_yea"] = cmbYea.Text;
            x["lot_sta"] = "";
            
            x["lot_dot"] = "";
            x["lot_doy"] = "";
            x["lot_don"] = "";
            x["lot_dom"] = "";

            x["lot_dor"] = "";
            x["lot_art"] = "";

            t.Rows.Add(x);

            frmGesDocLottiDettaglio f = new frmGesDocLottiDettaglio();
            f._rowLot = x;
            f.ShowDialog();

            _clsFun.PosRowLast(dgv1);
        }

        private void btnMdy_Click(object sender, EventArgs e)
        {
            Dettaglio();
        }

        private void Dettaglio()
        {
            CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, this.dgv1.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {
                DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                DataRow x = r.Row;

                frmGesDocLottiDettaglio f = new frmGesDocLottiDettaglio();
                f._rowLot = x;
                f._strArt = _strArt;
                f.ShowDialog();

                x = f._rowLot;
            }
        }

        private void cmbSta_SelectionChangeCommitted(object sender, EventArgs e)
        {
            FillDati();
        }

        private void cmbYea_SelectionChangeCommitted(object sender, EventArgs e)
        {
            FillDati();
        }

        private Boolean Salva()
        {
            Boolean b = true;

            if (dgv1.CurrentCell is DataGridViewCell)
            {
                dgv1.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }

            string s = "";
            DataRow x;
            DataRow[] j;

            s = "SELECT * FROM GesDocLotti";
            DataTable tLot = _clsFun.FillTabSql(TABDOCLOT, s, false, _strConSql);

            DataTable t = (DataTable)dgv1.DataSource;

            ArrayList aWhe = new ArrayList();
            aWhe.Add("lot_yea");
            aWhe.Add("lot_for");
            aWhe.Add("lot_ndo");
            aWhe.Add("lot_cod");
            aWhe.Add("lot_tip");
            ArrayList aExl = new ArrayList();

            foreach (DataRow y in t.Rows)
            {
                if ((string)y["LotMdy"] != "")
                {
                    x = tLot.NewRow();
                    x["lot_yea"] = y["lot_yea"];
                    x["lot_for"] = y["lot_for"];
                    x["lot_ndo"] = y["lot_ndo"];
                    x["lot_ddo"] = y["lot_ddo"];
                    x["lot_cod"] = y["lot_cod"];
                    x["lot_des"] = y["lot_des"];
                    x["lot_tip"] = y["lot_tip"];
                    x["lot_ann"] = y["lot_ann"];
                    x["lot_nat"] = y["lot_nat"];
                    x["lot_all"] = y["lot_all"];
                    x["lot_mac"] = y["lot_mac"];
                    x["lot_sez"] = y["lot_sez"];

                    x["lot_lot"] = y["lot_lot"];
                    x["lot_mna"] = y["lot_mna"];
                    x["lot_mdt"] = y["lot_mdt"];
                    x["lot_sbo"] = y["lot_sbo"];
                    x["lot_sta"] = y["lot_sta"];

                    //if(_strDoc != "")
                    //{
                    //    string[] a = _strDoc.Split('-');
                    //    if (a.Length > 4)
                    //    {
                    //        x["lot_dot"] = a[0];
                    //        x["lot_doy"] = a[1];
                    //        x["lot_don"] = a[2];
                    //        x["lot_dor"] = a[3];
                    //        x["lot_art"] = a[4];
                    //    }
                    //}

                    if (DBNull.Value.Equals(y["lot_idx"]) || (int)y["lot_idx"] == 0)
                    {
                        s = _clsFun.SqlInsertRow(TABDOCLOT, tLot, y);
                    }
                    else
                    {
                        aWhe = new ArrayList();
                        aWhe.Add("lot_idx");
                        s = "lot_idx=" + Convert.ToInt32(y["lot_idx"]);
                        j = tLot.Select(s);
                        if (j.Length > 0)
                            s = _clsFun.SqlUpdRow(TABDOCLOT, tLot, j[0], y, aWhe, aExl);
                    }

                    if (s != "")
                        _clsFun.SqlWrite(s, _strConSql);
                }
            }

            return b;
        }

        //private void btnCpy_Click(object sender, EventArgs e)
        //{
        //    LotCopia();
        //}

        private void LotCopia()
        {
            CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {
                DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                DataRow x = r.Row;

                DataTable t = (DataTable)dgv1.DataSource;

                DataRow y = t.NewRow();

                //y["lot_idx"] = x["lot_idx"];
                y["lot_cod"] = x["lot_cod"];
                y["lot_yea"] = x["lot_yea"];
                y["lot_tip"] = ""; // x["lot_tip"];
                y["lot_ndo"] = x["lot_ndo"];
                y["lot_ddo"] = x["lot_ddo"];
                y["lot_des"] = ""; // x["lot_des"];
                y["lot_for"] = x["lot_for"];
                y["ForDes"] = x["ForDes"];
                y["lot_ann"] = x["lot_ann"];

                y["lot_nat"] = x["lot_nat"];
                y["lot_all"] = x["lot_all"];
                y["lot_mac"] = x["lot_mac"];
                y["lot_sez"] = x["lot_sez"];

                y["lot_lot"] = x["lot_lot"];
                y["lot_mna"] = x["lot_mna"];
                y["lot_mdt"] = x["lot_mdt"];
                y["lot_sbo"] = x["lot_sbo"];
                y["lot_sta"] = x["lot_sta"];

                y["LotMdy"] = "";

                t.Rows.Add(y);
            }
        }

        //private void btnDel_Click(object sender, EventArgs e)
        //{
        //    CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
        //    if (cm.Position >= 0)
        //    {
        //        DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
        //        DataRow x = r.Row;

        //        string sLot = (string)x["lot_cod"];
        //        string sTip = (string)x["lot_tip"];
        //        string sIdx = Convert.ToString(x["lot_idx"]);

        //        if (MessageBox.Show("Confermi eliminazione del lotto " + sLot + " tipo " + sTip + "?", "ELIMINAZIONE LOTTO", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
        //        {
        //            string s = "DELETE FROM GesDocLotti WHERE lot_idx=" + sIdx;
        //            _clsFun.SqlWrite(s, _strConSql);
        //            x.Delete();
        //        }
        //    }
        //}

        private void dgv1_DoubleClick(object sender, EventArgs e)
        {
            if (_strSelect == "S")
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    _strYea = (string)x["lot_yea"];
                    if (_strTip == "RowDoc")
                        _strRes = (string)x["lot_cod"];
                    else
                        _strRes = ((string)x["lot_tip"]).PadRight(2, Convert.ToChar(' ')) + "-" + (string)x["lot_cod"] + "-" + (string)x["lot_yea"];

                    Esci();
                }
            }
            else
                Dettaglio();
        }

        private void txtSeek_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Enter)
                GrdSeek();
        }
        private void btnSeek_Click(object sender, EventArgs e)
        {
            GrdSeek();
        }

        private void GrdSeek()
        {
            string s = txtSeek.Text.Trim();

            string sFld = "TXT";
            if (_clsFun.Numerico(s))
                sFld = "NUM";

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

                        if (sFld == "NUM")
                        {
                            if (row.Cells["Lotto"].Value.ToString().ToLower().Contains(s.ToLower()) ||
                                row.Cells["N.documento"].Value.ToString().ToLower().Contains(s.ToLower()))
                            {
                                dgv1.Rows[row.Index].Selected = true;
                                dgv1.FirstDisplayedScrollingRowIndex = row.Index;
                                dgv1.CurrentCell = dgv1.Rows[row.Index].Cells[0];
                                b = true;
                                break;
                            }
                        }
                        else if (sFld == "TXT")
                        {
                            if (row.Cells["Descrizione"].Value.ToString().ToLower().Contains(s.ToLower()) ||
                            row.Cells["Fornitore"].Value.ToString().ToLower().Contains(s.ToLower()))
                            {
                                dgv1.Rows[row.Index].Selected = true;
                                dgv1.FirstDisplayedScrollingRowIndex = row.Index;
                                dgv1.CurrentCell = dgv1.Rows[row.Index].Cells[0];
                                b = true;
                                break;
                            }
                        }
                    }
                }
                if (!b)
                    dgv1.CurrentCell = dgv1.Rows[0].Cells[0];
            }

        }

        private void txtSeek_DoubleClick(object sender, EventArgs e)
        {
            TextBox txt = (TextBox)sender;

            frmUtyTastiera f = new frmUtyTastiera();
            f._strStr = txt.Text;
            f.ShowDialog();
            txt.Text = f._strStr;
        }

        private void stampaElencoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PdfLottiElenco();
        }

        private void stampaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PdfLottiControllo();
        }

        private void PdfLottiControllo()
        {
            string s = "";
            DataRow[] j;
            int iRows = 37;
            int iRow = 0;

            DataTable t = (DataTable)dgv1.DataSource;

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
                    s = DateTime.Today.ToString("dd/MM/yyyy") + " STAMPA ELENCO LOTTI";
                    gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    X += 20;

                    s = "Lotto";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 60, X, XStringFormats.Default);

                    s = "Tipo";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 100, X, XStringFormats.Default);

                    s = "Fornitore";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 150, X, XStringFormats.Default);

                    s = "Documento";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 300, X, XStringFormats.Default);

                    s = "Data";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 400, X, XStringFormats.Default);

                    //s = "Descrizione";
                    //gfx.DrawString(s, font1, XBrushes.Black, Y + 450, X, XStringFormats.Default);

                    X += 15;
                }

                iRow++;

                if (t.Rows.Count - 1 >= i)
                {
                    s = (string)t.Rows[i]["lot_cod"];
                    gfx.DrawString(s, font4, XBrushes.Black, Y + 60, X, XStringFormats.Default);

                    s = (string)t.Rows[i]["lot_tip"];
                    gfx.DrawString(s, font4, XBrushes.Black, Y + 100, X, XStringFormats.Default);

                    s = (string)t.Rows[i]["ForDes"];
                    gfx.DrawString(s, font4, XBrushes.Black, Y + 150, X, XStringFormats.Default);

                    s = (string)t.Rows[i]["lot_ndo"];
                    gfx.DrawString(s, font4, XBrushes.Black, Y + 350, X + 3, frmDX);

                    s = ((DateTime)t.Rows[i]["lot_ddo"]).ToString("dd/MM/yyyy");
                    gfx.DrawString(s, font4, XBrushes.Black, Y + 430, X + 3, frmDX);

                    s = (string)t.Rows[i]["lot_des"];
                    gfx.DrawString(s, font4, XBrushes.Black, Y + 480, X + 15, frmDX);

                    //break;

                    dRig++;
                }

                X += 40;
                //break;
            }

            X += 25;

            //s = "Totale righe " + lblInvRig.Text;
            //gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

            // Save the document...
            try
            {
                string sFil = "C:\\ApProject\\PDF\\Lot_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                pd.Save(sFil);
                // ...and start a viewer.
                Process.Start(sFil);
            }
            catch (Exception ex)
            {
                MessageBox.Show("File di stampa già aperto!");
            }
        }

        private void PdfLottiElenco()
        {
            string s = "";
            DataRow[] j;
            DataTable tTmp = new DataTable();
            DataTable tLot = (DataTable)dgv1.DataSource;

            DataTable t = tLot.Copy();

            s = _clsFun.ParGet(clsDefine.enuParametri.Par031Lotti2Pos, _strConSql);
            string sPrf = "";
            string[] a = s.Split(',');
            if (a.Length > 1 && a[1] != "")
                sPrf = a[1];

            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "LotQtc",
                Caption = "Qta caricata",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "LotQtu",
                Caption = "Qta uscita",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "LotQti",
                Caption = "Qta %",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });

            foreach(DataRow y in t.Rows)
            {
                s = (string)y["lot_des"];
                if ((string)y["lot_cod"] == "00008")
                    Console.WriteLine("zzzzzzzzz");

                tTmp = _clsQry.LotCarico(y);

                if (tTmp.Rows.Count > 0)
                {
                    y["LotQtc"] = (decimal)tTmp.Rows[0]["mov_qkg"];
                    y["LotQtu"] = _clsQry.LotVenduto(tTmp, sPrf + Convert.ToInt32(y["lot_cod"]).ToString("000"));
                    if ((decimal)y["LotQtu"] > 0 && (decimal)y["LotQtc"] > 0)
                        y["LotQti"] = ((decimal)y["LotQtu"] / (decimal)y["LotQtc"] * 100).ToString("#0.00");
                }
            }

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Lotti";

            // Create an empty page
            PdfPage page = pd.AddPage();

            page.Height = 842.0;
            page.Width = 595;
            page.Orientation = PdfSharp.PageOrientation.Landscape;

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 12, XFontStyle.Bold);
            XFont font2 = new XFont("Courier new", 9, XFontStyle.Bold);
            XFont font3 = new XFont("Courier new", 8, XFontStyle.Regular);
            XFont font4 = new XFont("Courier new", 10, XFontStyle.Bold);

            double Y = 5;
            double X = 20;

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            //string sFam = ""; // (string)t.Rows[0]["InvL1c"] + (string)t.Rows[0]["InvL2c"] + (string)t.Rows[0]["InvL3c"];
            int iRows = 18;
            int iRow = 0;
            int iPag = 0;

            decimal dRig = 0;
            //decimal dPez = 0;
            //decimal dVal = 0;

            XPen pen = new XPen(XColors.Black, 1.2);

            for (int i = 0; i <= t.Rows.Count - 1; i++)
            {
                if (iRow > iRows || iRow == 0)
                {
                    if (iRow > 0)
                    {
                        iPag++;
                        PdfLottiControlloPiede(iPag, gfx, pen, font2, frmDX);

                        page = pd.AddPage();
                        page.Height = 842.0;
                        page.Width = 595;
                        page.Orientation = PdfSharp.PageOrientation.Landscape;

                        gfx = XGraphics.FromPdfPage(page);
                    }

                    X = 20;
                    iRow = 0;

                    gfx.DrawRectangle(pen, Y + 5, X - 10, 820, 20);

                    s = DateTime.Today.ToString("dd/MM/yyyy") + new string(' ', 20) + " TRACCIABILITA' LOTTI APERTI";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 10, X + 3, XStringFormats.Default);

                    X += 40;

                    gfx.DrawRectangle(pen, Y + 5, X - 15, 820, 30);

                    s = "Lotto";
                    gfx.DrawString(s, font2, XBrushes.Black, Y + 10, X, XStringFormats.Default);

                    s = "Data carico";
                    gfx.DrawString(s, font2, XBrushes.Black, Y + 70, X, XStringFormats.Default);

                    s = "Ragione sociale";
                    gfx.DrawString(s, font2, XBrushes.Black, Y + 150, X, XStringFormats.Default);

                    s = "Descrizione";
                    gfx.DrawString(s, font2, XBrushes.Black, Y + 330, X, XStringFormats.Default);

                    s = "Lotto";
                    gfx.DrawString(s, font2, XBrushes.Black, Y + 530, X, XStringFormats.Default);
                    s = "fornitore";
                    gfx.DrawString(s, font2, XBrushes.Black, Y + 530, X + 10, XStringFormats.Default);

                    s = "Q.tà";
                    gfx.DrawString(s, font2, XBrushes.Black, Y + 660, X, XStringFormats.Default);
                    s = "carico";
                    gfx.DrawString(s, font2, XBrushes.Black, Y + 660, X+10, XStringFormats.Default);

                    s = "Q.tà";
                    gfx.DrawString(s, font2, XBrushes.Black, Y + 708, X, XStringFormats.Default);
                    s = "venduta";
                    gfx.DrawString(s, font2, XBrushes.Black, Y + 708, X+10, XStringFormats.Default);

                    s = "%";
                    gfx.DrawString(s, font2, XBrushes.Black, Y + 780, X, XStringFormats.Default);
                    s = "venduta";
                    gfx.DrawString(s, font2, XBrushes.Black, Y + 770, X+10, XStringFormats.Default);

                    X += 30;
                }

                iRow++;

                if (t.Rows.Count - 1 >= i)
                {
                    s = ((string)t.Rows[i]["lot_yea"]).Substring(2) + ((string)t.Rows[i]["lot_cod"]).PadLeft(4,Convert.ToChar("0"));
                    gfx.DrawString(s, font4, XBrushes.Black, Y + 10, X, XStringFormats.Default);

                    s = ((DateTime)t.Rows[i]["lot_ddo"]).ToString("dd/MM/yyyy");
                    gfx.DrawString(s, font4, XBrushes.Black, Y + 70, X, XStringFormats.Default);

                    if (!DBNull.Value.Equals(t.Rows[i]["ForDes"]))
                    {
                        s = ((string)t.Rows[i]["ForDes"] + new string(' ', 50)).Substring(0, 50);
                        gfx.DrawString(s, font4, XBrushes.Black, Y + 150, X, XStringFormats.Default);
                    }

                    s = ((string)t.Rows[i]["lot_des"] + new string(' ', 50)).Substring(0,30);
                    gfx.DrawString(s, font4, XBrushes.Black, Y + 330, X, XStringFormats.Default);

                    s = (string)t.Rows[i]["lot_lot"];
                    gfx.DrawString(s, font4, XBrushes.Black, Y + 530, X, XStringFormats.Default);

                    s = ((Decimal)t.Rows[i]["LotQtc"]).ToString("#0.00");
                    gfx.DrawString(s, font4, XBrushes.Black, Y + 695, X + 3, frmDX);

                    s = ((Decimal)t.Rows[i]["LotQtu"]).ToString("#0.00");
                    gfx.DrawString(s, font4, XBrushes.Black, Y + 750, X + 3, frmDX);

                    s = ((Decimal)t.Rows[i]["LotQti"]).ToString("#0.00");
                    gfx.DrawString(s, font4, XBrushes.Black, Y + 800, X + 3, frmDX);

                    //break;

                    dRig++;
                }

                X += 25;
                //break;
            }

            iPag++;
            PdfLottiControlloPiede(iPag, gfx, pen, font2, frmDX);

            //X += 25;
            //s = "Totale righe " + lblInvRig.Text;
            //gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

            // Save the document...
            try
            {
                string sFil = "C:\\ApProject\\PDF\\Lot_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                pd.Save(sFil);
                // ...and start a viewer.
                Process.Start(sFil);
            }
            catch (Exception ex)
            {
                MessageBox.Show("File di stampa già aperto!");
            }
        }

        private void PdfLottiControlloPiede(int intPag, XGraphics gfx, XPen pen, XFont font, XStringFormat frmDX)
        {
            gfx.DrawLine(pen, 830, 565, 10, 565);
            string s = "Pagina " + intPag.ToString();
            gfx.DrawString(s, font, XBrushes.Black, 820, 585, frmDX);
        }

        private void invioLOTTIABilanciaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Procedi con l'invio?", "LOTTI A BILANCIA", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                Lotti2Bizerba();
            }
        }

        private void Lotti2Bizerba()
        {
            string s = "";
            string sPath = Path.GetDirectoryName(_clsDef.TMPBILVAR) + "\\";
            string sFilLastLotto = sPath + "FilLastLotto.txt";                  //Leggo l'ultimo numero del lotto scritto per scriverlo in prima riga di winrint.dat
            string sFilwinrintBase = sPath + "winrint.bas";

            if (!File.Exists(sFilLastLotto))
                MessageBox.Show(sFilLastLotto + " non presente!", "CONTROLLO FILE ULTIMO LOTTO", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                DataTable tCnf = _clsQry.ConfSeek("", "POS");

                //s = (string)tCnb.Rows[0]["cnf_bil"];

                if ((string)tCnf.Rows[0]["cnf_bil"] == Convert.ToInt16((object)clsDefine.enuBilance.bilBizWinVarp).ToString("00"))
                {
                    s = (string)tCnf.Rows[0]["CnfVab"];

                    string[] a = s.Split(',');
                    if (a.Length == 0 || a[0].Trim() == "")
                        MessageBox.Show("Parametri bilancia non definiti!", "CONTROLLO BIZERBA", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    else
                    {
                        DataView v = new DataView((DataTable)dgv1.DataSource, "lot_sta='A' OR lot_sta='C'", "lot_cod", DataViewRowState.CurrentRows);

                        if (v.Count > 0)
                        {
                            string sFilDest = Path.GetDirectoryName(a[0]) + "\\winrint.dat";

                            string sFil = sPath + "winrint.dat";
                            //if (File.Exists(sFil))
                            //    File.Delete(sFil);

                            sFil = sPath + "winrint.dat";
                            if (File.Exists(sFil))
                                File.Delete(sFil);

                            //if (File.Exists(sPath + "winrint.bas"))
                            //    File.Copy(sPath + "winrint.bas", sFil);

                            Console.WriteLine("zzz");

                            s = "000";

                            using (StreamReader sr = new StreamReader(sFilLastLotto))
                            {
                                while (sr.Peek() >= 0)
                                {
                                    s = sr.ReadLine();
                                }
                                sr.Close();
                                sr.Dispose();
                            }

                            StreamWriter sw = new StreamWriter(sFil, true);

                            string sRig = "   " + s + new string(' ', 295);
                            sw.Write(sRig);

                            using (StreamReader sr = new StreamReader(sFilwinrintBase))
                            {
                                while (sr.Peek() >= 0)
                                {
                                    sRig = sr.ReadLine();
                                    sw.Write(sRig);
                                }
                                sr.Close();
                                sr.Dispose();
                            }

                            string sLot = "";

                            foreach (DataRowView y in v)
                            {
                                sLot = Convert.ToInt32(y["lot_cod"]).ToString("000");
                                sRig = "   " + sLot + new string(' ', 17);                          //Numero lotto                       1 - 20

                                s = ((string)y["lot_yea"]).Substring(2, 2) + Convert.ToInt32(y["lot_cod"]).ToString("0000");
                                sRig += s.PadRight(20, Convert.ToChar(' '));                        //Numero di identificazione         21 - 20
                                sRig += new string('0', 6);                                         //Campo non significativo           41 -  6
                                sRig += "1";                                                        //Abilitazione alla stampa dei dati 47 -  1
                                s = "0" + new string(' ', 14);
                                sRig += s;                                                          //Data Modifica (non significativo) 48 - 15
                                sRig += new string(' ', 20);                                        //Data di macellazione              63 - 20
                                sRig += ((string)y["lot_mna"]).PadRight(20, Convert.ToChar(' '));   //Paese di macellazione             83 - 20
                                sRig += ((string)y["lot_mac"]).PadRight(20, Convert.ToChar(' '));   //Nr. omologazione Azienda di Macellazione 103 -20
                                sRig += ((string)y["lot_all"]).PadRight(20, Convert.ToChar(' '));   //Paese Ingrasso                   123 - 20
                                sRig += ((string)y["lot_nat"]).PadRight(20, Convert.ToChar(' '));   //Paese di nascita                 143 - 20

                                s = ((string)y["lot_yea"]).Substring(2, 2) + Convert.ToInt32(y["lot_cod"]).ToString("0000");
                                sRig += s.PadRight(20, Convert.ToChar(' '));                        //Categoria                        163 - 20

                                sRig += ((string)y["lot_sez"]).PadRight(20, Convert.ToChar(' '));   //Paese di sezionamento            183 - 20
                                sRig += ((string)y["lot_sbo"]).PadRight(20, Convert.ToChar(' '));   //Nr. omologazione Azienda di Sezionamento 203 - 20
                                sRig += new string(' ', 20);                                        //Pezzo                            223 - 20
                                sRig += new string(' ', 20);                                        //Razza                            243 - 20
                                sRig += new string(' ', 20);                                        //Riserva                          263 - 20
                                sRig += new string(' ', 14);                                        //Riserva                          283 - 20
                                sRig += " ";                                                        //tipo record 0 = aggiungi/modifica 1 = cancella 303 - 1
                                sw.Write(sRig);
                                //sw.Write(sRig + "\r\n");
                            }

                            ((TextWriter)sw).Flush();
                            sw.Close();
                            sw.Dispose();

                            sw = new StreamWriter(sFilLastLotto, false);
                            sw.Write(sLot);
                            ((TextWriter)sw).Flush();
                            sw.Close();
                            sw.Dispose();


                            string sss = Path.GetDirectoryName(a[0]) + "\\winrint.dat";

                            Console.WriteLine("xxx");

                            if (File.Exists(sss))
                                File.Delete(sss);

                            File.Copy(sFil, sss);

                            s = sFil + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".dat";

                            File.Move(sFil, s);

                        }
                    }
                }
            }
        }

        private void cancellaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (_strSelect == "N")
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    string s = "";
                    string sYea = (string)x["lot_yea"];
                    string sFor = (string)x["lot_for"];
                    string sNdo = (string)x["lot_ndo"];
                    string sCod = (string)x["lot_cod"];
                    string sTip = (string)x["lot_tip"];
                    string sSta = (string)x["lot_sta"];

                    if (sSta != "C")
                        MessageBox.Show("Il Lotto deve essere in stato CANCELLATO!", "CANCELLAZIONE LOTTO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    else
                    {
                        string sMsg = "Confermi la cancellazione del LOTTO " + sCod + _clsDef.CRLF;
                        sMsg += "del fornitore " + sFor + _clsDef.CRLF;
                        sMsg += "dell'anno " + sYea + _clsDef.CRLF;
                        sMsg += "del documento " + sNdo + "?";

                        if (MessageBox.Show(sMsg, "CANCELLAZIONE LOTTO", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                        {

                            s = "DELETE FROM GesDocLotti WHERE ";
                            s += "lot_yea='" + sYea + "' AND ";
                            s += "lot_for='" + sFor + "' AND ";
                            s += "lot_ndo='" + sNdo + "' AND ";
                            s += "lot_cod='" + sCod + "' AND ";
                            s += "lot_tip='" + sTip + "'";

                            _clsFun.SqlWrite(s, _strConSql);

                            x.Delete();

                        }
                    }

                }
            }

        }
    }
}
