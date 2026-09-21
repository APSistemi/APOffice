using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System.IO;

namespace APOffice
{
    public partial class frmPrnPlu : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        private const string TABANAART = "AnaArticoli";

        private string _strConSql = "";

        public frmPrnPlu()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0); ;
        }

        private void frmPrnPlu_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");
            SetDgv1();
            FillTab();

            DataTable t = new clsGenTabTmp().TabTmpArtPluBilance("TabPlu");
            dgv1.DataSource = t;
        }
 
        private void frmPrnPlu_KeyDown(object sender, KeyEventArgs e)
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
            this.Close();
        }

        private void cmbArtReb_SelectionChangeCommitted(object sender, EventArgs e)
        {
            cmbArtReb.Enabled = false;
            FillDati();
            cmbArtReb.Enabled = true;
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
            cTbc.DataPropertyName = "tmp_plu";
            cTbc.Name = "PLU";
            cTbc.Width = 40;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "##0.00";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_art";
            cTbc.Name = "Articolo";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_des";
            cTbc.Name = "Descrizione";
            cTbc.Width = 300;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_cos";
            cTbc.Name = "Costo";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "##0.00";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_pve";
            cTbc.Name = "Prezzo";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc); 
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "##0.00";

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_sta";
            cTbc.Name = "Stato";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_ean";
            cTbc.Name = "Barcode";
            cTbc.Width = 90;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_tas";
            cTbc.Name = "Tasto";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "tmp_red";
            cTbc.Name = "Reparto";
            cTbc.Width = 150;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);
        }

        private void FillTab()
        {
            string p = "TabRepBilance";
            string s = "SELECT * FROM TabRepBilance WHERE tab_ann=0";
            DataTable t = _clsFun.FillTabSql(p, s, false, _strConSql);
            DataRow  x = t.NewRow();
            x["tab_cod"] = "";
            x["tab_des"] = "  Non definito";
            t.Rows.InsertAt(x, 0);
            cmbArtReb.DataSource = t;
            cmbArtReb.DisplayMember = "tab_des";
            cmbArtReb.ValueMember = "tab_cod";
            cmbArtReb.SelectedValue = "";
        }

        private void FillDati()
        {
            DataRow x;
            DataRow[] j;
            DataTable tTmp = new DataTable();
            DataTable t = (DataTable)dgv1.DataSource;
            t.Clear();

            DataColumn[] keys = new DataColumn[1];
            keys[0] = t.Columns["tmp_plu"];
            keys[0] = t.Columns["tmp_art"];
            t.PrimaryKey = keys;

            if (cmbArtReb.SelectedValue.ToString() != "")
            {
                string s = "SELECT ";
                s += "art_cod, art_des, art_plu, art_sta, tab_des AS RepDes ";
                s += "FROM AnaArticoli ";
                s += "LEFT OUTER JOIN TabReparti ON art_rep = TabReparti.tab_cod ";
                s += "WHERE ";
                s += "art_reb='" + cmbArtReb.SelectedValue.ToString() + "' ";
                if (chkArtAtt.Checked)
                    s += " AND art_sta='" + _clsDef.STAATT + "'";
                s += "ORDER BY art_plu";

                DataTable tArt = _clsFun.FillTabSql(TABANAART, s, false, _strConSql);

                progressBar1.Value = 0;
                progressBar1.Maximum = tArt.Rows.Count;
                progressBar1.Minimum = 0;

                foreach (DataRow y in tArt.Rows)
                {
                    progressBar1.Increment(1);
                    Application.DoEvents();

                    if ((string)y["art_cod"] == "0013270")
                        Console.WriteLine("aaaaaaaaaaa");

                    tTmp = _clsQry.SeekArtEan((string)y["art_cod"], "");
                    if (tTmp.Rows.Count > 0)
                    {
                        x = t.NewRow();
                        x["tmp_art"] = tTmp.Rows[0]["tmp_art"];
                        x["tmp_des"] = tTmp.Rows[0]["tmp_ard"];

                        x["tmp_plu"] = tTmp.Rows[0]["tmp_plu"];
                        if (_clsFun.Numerico(tTmp.Rows[0]["tmp_plu"]))
                            x["tmp_plu"] = Convert.ToInt32(tTmp.Rows[0]["tmp_plu"]).ToString().PadLeft(6, Convert.ToChar(" "));

                        x["tmp_tas"] = tTmp.Rows[0]["tmp_tas"];

                        x["tmp_cos"] = tTmp.Rows[0]["tmp_cos"];
                        x["tmp_pve"] = tTmp.Rows[0]["tmp_prv"];
                        x["tmp_sta"] = tTmp.Rows[0]["tmp_sta"];

                        if ((Boolean)tTmp.Rows[0]["tmp_bil"])
                            x["tmp_ean"] = tTmp.Rows[0]["tmp_ean"];
                        else if (tTmp.Rows.Count > 1)
                        {
                            foreach (DataRow yy in tTmp.Rows)
                            {
                                if ((Boolean)yy["tmp_bil"])
                                    x["tmp_ean"] = yy["tmp_ean"];
                            }
                        }

                        x["tmp_rep"] = tTmp.Rows[0]["tmp_rep"];
                        x["tmp_red"] = y["RepDes"];


                        t.Rows.Add(x);
                    }
                    else if ((Boolean)tTmp.Rows[0]["tmp_bil"])
                    {
                        j = t.Select("tmp_art='" + y["tmp_art"] + "'");
                        if (j.Length > 0)
                            j[0]["tmp_ean"] = tTmp.Rows[0]["tmp_ean"];
                    }

                }

                DataView v = new DataView(t, "", "tmp_plu", DataViewRowState.CurrentRows);

                dgv1.DataSource = v.ToTable();

                lblCnt.Text = t.Rows.Count.ToString();
            }
        }

        private void dgv1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 2)
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    string s = (string)x["tmp_art"];
                    frmAnaArticolo f = new frmAnaArticolo();
                    f._strArtCod = s;
                    f.ShowDialog();
                }
            }

        }

        private void dgv1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 2)
            {
                CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                if (cm.Position >= 0)
                {
                    DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                    DataRow x = r.Row;

                    string s = (string)x["tmp_art"];
                    frmAnaArticolo f = new frmAnaArticolo();
                    f._strArtCod = s;
                    f.ShowDialog();
                }
            }
        }

        private void stampaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PrnPdfPlu();
        }

        private void stampaConIngredientiToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PrnPdfPluIngredienti();
        }

        public void PrnPdfPlu()
        {
            string s = "";
            int iRows = 37;
            int iRow = 0;
            decimal d = 0;

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "PLU";

            // Create an empty page
            PdfPage page = pd.AddPage();

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 11, XFontStyle.Bold);
            XFont font2 = new XFont("Courier new", 18, XFontStyle.Regular);
            XFont font3 = new XFont("Courier new", 8, XFontStyle.Regular);

            double Y = 5;
            double X = 20;

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            DataTable t = (DataTable)dgv1.DataSource;

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
                    s = DateTime.Now.ToString("dd/MM/yyyy") + " STAMPA PLU BILANCIA REPARTO " + cmbArtReb.Text;
                    //if (chkCli.Checked)
                    //    s += "CLIENTI ";
                    //if (chkFor.Checked)
                    //    s += "FORNITORI ";
                    gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    X += 30;

                    s = "PLU";
                    gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    //s = "Tasto";
                    //gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    s = "Articolo";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 40, X, XStringFormats.Default);

                    s = "Descrizione";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 100, X, XStringFormats.Default);

                    s = "Costo";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 420, X, XStringFormats.Default);

                    s = "Vendita";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 480, X, XStringFormats.Default);

                    s = "Stato";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 550, X, XStringFormats.Default);

                    X += 35;
                }

                iRow++;

                if (t.Rows.Count - 1 >= i)
                {
                    s = (string)t.Rows[i]["tmp_plu"];
                    gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    //s = (string)t.Rows[i]["tmp_tas"];
                    //gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    s = (string)t.Rows[i]["tmp_art"];
                    gfx.DrawString(s, font1, XBrushes.Black, Y+48, X, XStringFormats.Default);

                    s = (string)t.Rows[i]["tmp_des"];
                    gfx.DrawString(s, font1, XBrushes.Black, Y+100, X, XStringFormats.Default);

                    s = ((decimal)t.Rows[i]["tmp_cos"]).ToString("####,##0.00");
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 462, X + 3, frmDX);

                    s = ((decimal)t.Rows[i]["tmp_pve"]).ToString("####,##0.00");
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 530, X + 3, frmDX);

                    s = (string)t.Rows[i]["tmp_sta"];
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 550, X, XStringFormats.Default);

                    //d += (decimal)t.Rows[i]["vre_ven"];
                    //break;
                }

                X += 20;
                //break;

            }

            X += 15;

            //s = "Totale    " + d.ToString();
            //gfx.DrawString(s, font1, XBrushes.Black, Y + 240, X, XStringFormats.Default);

            // Save the document...
            try
            {
                string sFil = "C:\\ApProject\\PDF\\PluBil_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                pd.Save(sFil);
                // ...and start a viewer.
                Process.Start(sFil);
            }
            catch (Exception ex)
            {
                //MessageBox.Show("File di stampa già aperto!");
                Console.WriteLine("");
            }

        }

        private void excelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ArticoliPluCsv();
        }

        private void ArticoliPluCsv()
        {
            string s = "";
            DataRow[] j;
            
            DataView v = new DataView((DataTable)dgv1.DataSource, "", "tmp_plu", DataViewRowState.CurrentRows);

            DataTable t = ((DataTable)dgv1.DataSource).Clone();

            //t = v.Table.Clone();

            //string s = "";
            //decimal d = 0;

            foreach (DataRowView r in v)
            {
                r["tmp_plu"] = ((string)r["tmp_plu"]).Trim().PadLeft(6, Convert.ToChar(" "));
                r["tmp_tas"] = ((string)r["tmp_tas"]).Trim().PadLeft(4, Convert.ToChar(" "));
                t.ImportRow(r.Row);
            }

            //v = new DataView(t, "", "tmp_plu", DataViewRowState.CurrentRows);

            //DataTable t2 = v.ToTable();


            string sTit = DateTime.Today.ToString("dd/MM/yyyy") + " - Estrazione PLU " + cmbArtReb.Text;
            string sFil = "";

            s = "";
            s += "tmp_plu, PLU, 60, StringLiteral;";
            s += "tmp_art, articolo, 50, StringLiteral;";
            s += "tmp_des, Descrizione, 50, StringLiteral;";
            s += "tmp_cos, Costo, 50, StringLiteral;";
            s += "tmp_pve, P.Vendita, 70, StringLiteral;";
            s += "tmp_sta, Stato, 50, Decimal2;";
            s += "tmp_ean, Barcode, 50, Integer;";
            s += "tmp_tas, Tasto, 50, Integer;";

            string sFld = s;

            (new clsExcel()).exportToCsv1(sFld, t, sFil, sTit, "");
        }

        public void PrnPdfPluIngredienti()
        {
            string s = "";
            decimal iRows = 37;
            decimal iRow = 0;
            decimal d = 0;

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "PLU";

            // Create an empty page
            PdfPage page = pd.AddPage();

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 11, XFontStyle.Bold);
            XFont font2 = new XFont("Courier new", 18, XFontStyle.Regular);
            XFont font3 = new XFont("Courier new", 8, XFontStyle.Regular);

            double Y = 5;
            double X = 20;

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            DataTable t = (DataTable)dgv1.DataSource;

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
                    s = DateTime.Now.ToString("dd/MM/yyyy") + " STAMPA PLU BILANCIA REPARTO " + cmbArtReb.Text;
                    //if (chkCli.Checked)
                    //    s += "CLIENTI ";
                    //if (chkFor.Checked)
                    //    s += "FORNITORI ";
                    gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    X += 30;

                    s = "PLU";
                    gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    //s = "Tasto";
                    //gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    s = "Articolo";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 40, X, XStringFormats.Default);

                    s = "Descrizione";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 100, X, XStringFormats.Default);

                    //s = "Costo";
                    //gfx.DrawString(s, font1, XBrushes.Black, Y + 420, X, XStringFormats.Default);

                    s = "Vendita";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 480, X, XStringFormats.Default);

                    //s = "Stato";
                    //gfx.DrawString(s, font1, XBrushes.Black, Y + 550, X, XStringFormats.Default);

                    X += 35;
                }

                iRow++;

                if (t.Rows.Count - 1 >= i)
                {
                    s = (string)t.Rows[i]["tmp_plu"];
                    gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    //s = (string)t.Rows[i]["tmp_tas"];
                    //gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    s = (string)t.Rows[i]["tmp_art"];
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 48, X, XStringFormats.Default);

                    s = (string)t.Rows[i]["tmp_des"];
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 100, X, XStringFormats.Default);

                    //s = ((decimal)t.Rows[i]["tmp_cos"]).ToString("####,##0.00");
                    //gfx.DrawString(s, font1, XBrushes.Black, Y + 462, X + 3, frmDX);

                    s = ((decimal)t.Rows[i]["tmp_pve"]).ToString("####,##0.00");
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 530, X + 3, frmDX);

                    //s = (string)t.Rows[i]["tmp_sta"];
                    //gfx.DrawString(s, font1, XBrushes.Black, Y + 550, X, XStringFormats.Default);

                    //d += (decimal)t.Rows[i]["vre_ven"];
                    //break;

                    if (i == 0)
                        Console.WriteLine("zzzzzzzzzz");

                    string sIng = "";
                    s = _clsDef.PATHINGREDIENTI + "et01_" + (string)t.Rows[i]["tmp_art"] + ".rtf";
                    if (File.Exists(s))
                    {
                        RichTextBox rch = new RichTextBox();
                        rch.LoadFile(s);
                        sIng = rch.Text;
                        //sIng = _clsFun.StrSplit(sIng, 40);

                        XTextFormatter tf = new XTextFormatter(gfx);

                        XRect rect = new XRect(Y, X+5, 400, 40);
                        gfx.DrawRectangle(XBrushes.Transparent, rect);
                        //tf.Alignment = ParagraphAlignment.Left; 
                        tf.DrawString(sIng, font3, XBrushes.Black, rect, XStringFormats.TopLeft);

                        if (sIng != "")
                        {
                            //string[] a = sIng.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);

                            //if(a.Length > 1 )
                            Console.WriteLine("zzzzzzzzzz");

                            int ii = Convert.ToInt16(sIng.Length / 80);

                            for (int i2 = 0; i2 <= ii; i2++)
                            {
                                X += 10;
                                iRow += Convert.ToDecimal(0.5);

                            }
                        }
                    }

                    XPen pen = new XPen(XColors.Black, 0.5);
                    gfx.DrawLine(pen, Y, X + 5, Y + 600, X + 5);
                }

                X += 20;
                //break;

            }

            X += 15;

            //s = "Totale    " + d.ToString();
            //gfx.DrawString(s, font1, XBrushes.Black, Y + 240, X, XStringFormats.Default);

            // Save the document...
            try
            {
                string sFil = "C:\\ApProject\\PDF\\PluBil_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                pd.Save(sFil);
                // ...and start a viewer.
                Process.Start(sFil);
            }
            catch (Exception ex)
            {
                //MessageBox.Show("File di stampa già aperto!");
                Console.WriteLine("");
            }

        }



    }
}
