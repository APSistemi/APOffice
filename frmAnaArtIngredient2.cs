using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Text.RegularExpressions;

namespace APOffice
{
    public partial class frmAnaArtIngredient2 : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        private const string TABARTING = "AnaArtIngredienti2";

        private string _strConSql = "";

        public string _strArtCod = "";
        //public string _strDes = "";
        //public string _strEan = "";
        //public decimal _decPes = 0;
        //public decimal _decTar = 0;
        //public decimal _decPrv = 0;
        //public decimal _decImp = 0;

        public frmAnaArtIngredient2()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
        }

        private void frmAnaArtIngredienti_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");

            lstColors.DropDown.Items.Add("Black");
            lstColors.DropDown.Items.Add("Red");
            lstColors.DropDown.Items.Add("Blue");

            //InstalledFontCollection fonts = new InstalledFontCollection();
            //foreach (FontFamily family in fonts.Families)
            //{
            //    lstFonts.DropDown.Items.Add(family.Name);
            //}

            lstFonts.DropDown.Items.Add("Arial");
            lstFonts.DropDown.Items.Add("Arial Black");
            lstFonts.DropDown.Items.Add("Arial Narrow");
            lstFonts.DropDown.Items.Add("Arial Rounded MT Bold");
            lstFonts.DropDown.Items.Add("Times New Roman");
            lstFonts.DropDown.Items.Add("Verdana");

            lstZoom.DropDown.Items.Add("300%");
            lstZoom.DropDown.Items.Add("200%");
            lstZoom.DropDown.Items.Add("100%");

            lstFontSize.DropDown.Items.Add("8");
            lstFontSize.DropDown.Items.Add("10");
            lstFontSize.DropDown.Items.Add("12");
            lstFontSize.DropDown.Items.Add("15");
            lstFontSize.DropDown.Items.Add("20");

            FillDati();
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }
        private void frmAnaArtIngredienti_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
        }
        private void Esci()
        {
            Salva();
            this.Close();
        }

        private void nuovaRigaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //DataTable t = (DataTable)dgv1.DataSource;

            //string sRow = "001";
            //string sTip = "";

            //if (t.Rows.Count > 0)
            //{
            //    DataView v = new DataView(t, "", "ing_row DESC", DataViewRowState.CurrentRows);
            //    if (v.Count > 0)
            //    {
            //        sRow = (Convert.ToDecimal(v[0]["ing_row"]) + 1).ToString("000");
            //        sTip = (string)v[0]["ing_tip"];
            //    }
            //}

            //DataRow x = t.NewRow();
            //x["ing_art"] = _strArtCod;
            //x["ing_tip"] = sTip;
            //x["ing_crt"] = "";
            //x["ing_row"] = sRow;
            //x["ing_txt"] = "";
            //x["ing_ann"] = false;
            //t.Rows.Add(x);
        }

        private void FillDati()
        {
            string s = "";

            DataTable t = _clsQry.SeekArtEan(_strArtCod, "");
            if (t.Rows.Count > 0)
            {
                DataRow r0 = t.Rows[0];
                string sDes = r0["tmp_ard"] != DBNull.Value ? Convert.ToString(r0["tmp_ard"]) : "";

                decimal dPes = 0m;
                if (r0["tmp_pne"] != DBNull.Value)
                    decimal.TryParse(Convert.ToString(r0["tmp_pne"]).Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out dPes);

                decimal dTar = 0m;
                if (r0["tmp_tar"] != DBNull.Value)
                    decimal.TryParse(Convert.ToString(r0["tmp_tar"]).Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out dTar);

                decimal dPrv = 0m;
                if (r0["tmp_prv"] != DBNull.Value)
                    decimal.TryParse(Convert.ToString(r0["tmp_prv"]).Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out dPrv);

                decimal dImp = dPrv * dPes;

                double dGsc = 0;
                if (r0["tmp_gsc"] != DBNull.Value)
                {
                    string sGscRaw = Convert.ToString(r0["tmp_gsc"]).Trim();
                    if (!string.IsNullOrEmpty(sGscRaw))
                    {
                        double.TryParse(sGscRaw.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out dGsc);
                    }
                }

                string sEan = r0["tmp_ean"] != DBNull.Value ? Convert.ToString(r0["tmp_ean"]).Trim() : "";

                if (dTar > 0)
                    dTar = dTar / 1000;

                decimal dVal = 0;
                if (dPes > 0 && dPrv > 0)
                    dVal = (dPes - dTar) * dPrv;

                string sPve = "0000";
                if (dVal > 0)
                    sPve = (dVal * 100).ToString("0000");

                if (sEan != "")
                {
                    sEan = sEan.Substring(0, 8) + sPve;
                    sEan = sEan + new clsCtrlCodici().FindMod10Digit(sEan);
                }

                lblDes.Text = sDes;
                lblPes.Text = dPes.ToString("#0.00");
                lblTar.Text = dTar.ToString("#0.000");
                lblPrv.Text = dPrv.ToString("#0.00");
                lblImp.Text = dImp.ToString("#0.00");
                lblEan.Text = sEan;

                try
                {
                    dtpDsc.Value = DateTime.Now.AddDays(dGsc);
                }
                catch
                {
                    dtpDsc.Value = DateTime.Today;
                }
            }

            //s = "";
            //s += "SELECT * FROM AnaArtIngredienti2 WHERE ing_art='" + _strArtCod + "'";
            //t = _clsFun.FillTabSql(TABARTING, s, false, _strConSql);
            //if(t.Rows.Count > 0)
            //{
            //    if(!DBNull.Value.Equals(t.Rows[0]["ing_img"]))
            //    {
            //        Byte[] img = (Byte[])(t.Rows[0]["ing_img"]);
            //        Byte[] rtf = new Byte[Convert.ToInt32((img.GetBytes(0, 0, null, 0, Int32.MaxValue)))];
            //        long bytesReceived = reader.GetBytes(0, 0, rtf, 0, rtf.Length);
            //        ASCIIEncoding encoding = new ASCIIEncoding();
            //        rchImg.Rtf = encoding.GetString(rtf, 0, Convert.ToInt32(bytesReceived));
            //    }
            //    //rchTx2.Rtf = Convert.ToString(t.Rows[0]["ing_tx2"]);
            //}

            //SqlConnection cn = new SqlConnection(_strConSql);
            //cn.Open();
            //SqlCommand cmd = new SqlCommand("SELECT ing_img FROM AnaArtIngredienti2 WHERE ing_art='" + _strArtCod + "'", cn);
            //SqlDataReader rdr = cmd.ExecuteReader();
            //rdr.Read();
            //if (rdr.HasRows)
            //{
            //    if (!rdr.IsDBNull(0))
            //    {
            //        Byte[] rtf = new Byte[Convert.ToInt32((rdr.GetBytes(0, 0, null, 0, Int32.MaxValue)))];
            //        long bytesReceived = rdr.GetBytes(0, 0, rtf, 0, rtf.Length);

            //        ASCIIEncoding encoding = new ASCIIEncoding();
            //        rchImg.Rtf = encoding.GetString(rtf, 0, Convert.ToInt32(bytesReceived));
            //    }
            //}

            //dgv1.DataSource = t;

            string sFil = _clsDef.PATHINGREDIENTI + "et01_" + _strArtCod + ".rtf";

            if (File.Exists(sFil))
                rchImg.LoadFile(sFil);
        }

        private void oldFillDati()
        {

            rchImg.Clear();

            SqlConnection cn = null;
            SqlCommand cmd = null;
            SqlDataReader reader = null;
            try
            {
                //cn = new SqlConnection("Database=Northwind;Integrated Security=true;");
                cn = new SqlConnection(_strConSql);

                cn.Open();
                cmd = new SqlCommand("SELECT ing_img FROM AnaArtIngredienti2 WHERE ing_art='0000001'", cn);
                reader = cmd.ExecuteReader();
                reader.Read();
                if (reader.HasRows)
                {
                    if (!reader.IsDBNull(0))
                    {
                        Byte[] rtf = new Byte[Convert.ToInt32((reader.GetBytes(0, 0, null, 0, Int32.MaxValue)))];
                        long bytesReceived = reader.GetBytes(0, 0, rtf, 0, rtf.Length);

                        ASCIIEncoding encoding = new ASCIIEncoding();
                        rchImg.Rtf = encoding.GetString(rtf, 0, Convert.ToInt32(bytesReceived));
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (null != reader) reader.Close();
                if (null != cn) cn.Close();
            }
        }

        private void okSalva()
        {
            //if (dgv1.CurrentCell is DataGridViewCell)
            //{
            //    dgv1.CommitEdit(DataGridViewDataErrorContexts.Commit);
            //}

            string s = "";
            DataRow x;
            DataRow[] j;

            s = "SELECT * FROM AnaArtIngredienti2 WHERE ";
            s += "ing_art='" + _strArtCod + "' ";
            s += "ORDER BY ing_row";
            DataTable tIng = _clsFun.FillTabSql(TABARTING, s, false, _strConSql);

            //DataRow x = t.NewRow();
            //x["ing_art"] = _strArtCod;
            //x["ing_tip"] = sTip;
            //x["ing_crt"] = "";
            //x["ing_row"] = sRow;
            //x["ing_txt"] = "";
            //x["ing_ann"] = false;
            //t.Rows.Add(x);

            x = tIng.NewRow();
            x["ing_art"] = _strArtCod;
            x["ing_tip"] = "BIL";
            x["ing_crt"] = "";
            x["ing_row"] = "001";
            x["ing_txt"] = "";
            x["ing_ann"] = false;

            x["ing_tx2"] = rchImg.Rtf;

            ArrayList aWhe = new ArrayList();
            aWhe.Add("ing_art");
            aWhe.Add("ing_row");
            ArrayList aExl = new ArrayList();

            j = tIng.Select("ing_row='" + "001" + "'");
            if (j.Length > 0)
                s = _clsFun.SqlUpdRow(TABARTING, tIng, j[0], x, aWhe, aExl);
            else
                s = _clsFun.SqlInsertRow(TABARTING, tIng, x);

            if (s != "")
                _clsFun.SqlWrite(s, _strConSql);

            //DataTable t = (DataTable)dgv1.DataSource;

            //ArrayList aWhe = new ArrayList();
            //aWhe.Add("ing_art");
            //aWhe.Add("ing_row");
            //ArrayList aExl = new ArrayList();

            //foreach(DataRow y in t.Rows)
            //{
            //    x = tIng.NewRow();
            //    x["ing_art"] = y["ing_art"];
            //    x["ing_tip"] = y["ing_tip"];
            //    x["ing_crt"] = y["ing_crt"];
            //    x["ing_row"] = y["ing_row"];
            //    x["ing_txt"] = y["ing_txt"];
            //    x["ing_ann"] = y["ing_ann"];

            //    j = tIng.Select("ing_row='" + y["ing_row"] + "'");
            //    if (j.Length > 0)
            //        s = _clsFun.SqlUpdRow(TABARTING, tIng, j[0], y, aWhe, aExl);
            //    else
            //        s = _clsFun.SqlInsertRow(TABARTING, tIng, y);

            //    if(s != "")
            //        _clsFun.SqlWrite(s, _strConSql);
            //}
        }

        private void ookkSalva()
        {
            string s = "";
            FileStream stream = null;
            SqlConnection cn = null;
            SqlCommand cmd = null;
            try
            {
                rchImg.SaveFile("temp.rtf");
                stream = new FileStream("temp.rtf", FileMode.Open, FileAccess.Read);
                int size = Convert.ToInt32(stream.Length);
                Byte[] rtf = new Byte[size];
                stream.Read(rtf, 0, size);

                //cn = new SqlConnection("Database=Northwind;Integrated Security=true;");

                cn = new SqlConnection(_strConSql);
                cn.Open();

                s = "SELECT * FROM AnaArtIngredienti2 WHERE ing_art='" + _strArtCod + "'";
                DataTable tIng = _clsFun.FillTabSql(TABARTING, s, false, _strConSql);
                if(tIng.Rows.Count == 0)
                    cmd = new SqlCommand("INSERT INTO AnaArtIngredienti2 (ing_art, ing_img) VALUES ('" + _strArtCod + "',@Photo)", cn);
                else
                    cmd = new SqlCommand("UPDATE AnaArtIngredienti2 SET ing_img=@Photo WHERE ing_art='" + _strArtCod + "'", cn);

                SqlParameter paramRTF =
                    new SqlParameter("@Photo",
                                     SqlDbType.Image,
                                     rtf.Length,
                                     ParameterDirection.Input,
                                     false,
                                     0, 0, null,
                                     DataRowVersion.Current,
                                     rtf);
                cmd.Parameters.Add(paramRTF);

                int rowsUpdated = Convert.ToInt32(cmd.ExecuteNonQuery());
                //MessageBox.Show(String.Format("{0} rows updated", rowsUpdated));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (null != stream) stream.Close();
                if (null != cmd) cmd.Parameters.Clear();
                if (null != cn) cn.Close();
            }
        }

        private void Salva()
        {
            string s = "";

            string sFil = _clsDef.PATHINGREDIENTI + "et01_" + _strArtCod + ".rtf";

            FileStream stream = null;
            //SqlConnection cn = null;
            //SqlCommand cmd = null;
            try
            {
                //C:\ApProject\Temp\Ingredienti

                rchImg.SaveFile(sFil);
                //stream = new FileStream(sFil, FileMode.Open, FileAccess.Read);
                //int size = Convert.ToInt32(stream.Length);
                //Byte[] rtf = new Byte[size];
                //stream.Read(rtf, 0, size);

                //cn = new SqlConnection(_strConSql);
                //cn.Open();

                //s = "SELECT * FROM AnaArtIngredienti2 WHERE ing_art='" + _strArtCod + "'";
                //DataTable tIng = _clsFun.FillTabSql(TABARTING, s, false, _strConSql);
                //if (tIng.Rows.Count == 0)
                //    cmd = new SqlCommand("INSERT INTO AnaArtIngredienti2 (ing_art, ing_img) VALUES ('" + _strArtCod + "',@Photo)", cn);
                //else
                //    cmd = new SqlCommand("UPDATE AnaArtIngredienti2 SET ing_img=@Photo WHERE ing_art='" + _strArtCod + "'", cn);

                //SqlParameter paramRTF =
                //    new SqlParameter("@Photo",
                //                     SqlDbType.Image,
                //                     rtf.Length,
                //                     ParameterDirection.Input,
                //                     false,
                //                     0, 0, null,
                //                     DataRowVersion.Current,
                //                     rtf);
                //cmd.Parameters.Add(paramRTF);

                //int rowsUpdated = Convert.ToInt32(cmd.ExecuteNonQuery());
                ////MessageBox.Show(String.Format("{0} rows updated", rowsUpdated));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                if (null != stream) stream.Close();
                //if (null != cmd) cmd.Parameters.Clear();
                //if (null != cn) cn.Close();
            }
        }

        private void stampaEtichetteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //dgv1.CommitEdit(DataGridViewDataErrorContexts.Commit);

            //DataTable t = (DataTable)dgv1.DataSource;
            //DataView v = new DataView(t, "ing_tip='ETI' AND ing_ann=0", "ing_row", DataViewRowState.CurrentRows);

            //if( v.Count == 0)
            //    MessageBox.Show("Non sono presenti valori da stampare!", "CONTROLLO STAMPA", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //else if(!_clsFun.Numerico(txtQta.Text))
            //    MessageBox.Show("Quantità non definita correttamente!", "CONTROLLO STAMPA", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //else
            //{
            //    DataTable tTmp = v.ToTable();

            //    decimal dQta = Convert.ToDecimal(txtQta.Text);

            //    frmAnaArtIngLabel f = new frmAnaArtIngLabel();
            //    f._tabTab = tTmp.Copy();

            //    f._strDes = lblDes.Text;
            //    f._strPes = lblPes.Text.PadLeft(10,Convert.ToChar(' '));
            //    f._strTar = lblTar.Text.PadLeft(10, Convert.ToChar(' '));
            //    f._strPrv = lblPrv.Text.PadLeft(10, Convert.ToChar(' '));
            //    f._strImp = lblImp.Text.PadLeft(10, Convert.ToChar(' '));
            //    f._strEan = lblEan.Text.PadLeft(10, Convert.ToChar(' '));
            //    f._strDsc = dtpDsc.Value.ToShortDateString();

            //    f._decQta = dQta;
            //    f.ShowDialog();
            //}
        }

        private void rchImg_SelectionChanged(object sender, EventArgs e)
        {
            if (rchImg.SelectionFont != null)
            {
                cmdBold.Checked = rchImg.SelectionFont.Bold;
                cmdItalic.Checked = rchImg.SelectionFont.Italic;
                cmdUnderline.Checked = rchImg.SelectionFont.Underline;
                rchImg.SelectionFont = new Font("Arial", 12);
                //rchImg.SelectionColor = Color.Orange;
                //rchImg.SelectedText = "Oranges" + "\n";
                //rchImg.SelectionFont = new Font("Arial", 12);
                //rchImg.SelectionColor = Color.Purple;
                //rchImg.SelectedText = "Grapes" + "\n";



            }
        }
        
        private void cmdBold_Click(object sender, EventArgs e)
        {
            if (rchImg.SelectionFont == null)
            {
                return;
            }

            FontStyle style = rchImg.SelectionFont.Style;

            if (rchImg.SelectionFont.Bold)
            {
                style &= ~FontStyle.Bold;
            }
            else
            {
                style |= FontStyle.Bold;

            }
            rchImg.SelectionFont = new Font(rchImg.SelectionFont, style);
        }
        private void cmdItalic_Click(object sender, EventArgs e)
        {
            if (rchImg.SelectionFont == null)
            {
                return;
            }
            FontStyle style = rchImg.SelectionFont.Style;

            if (rchImg.SelectionFont.Italic)
            {
                style &= ~FontStyle.Italic;
            }
            else
            {
                style |= FontStyle.Italic;
            }
            rchImg.SelectionFont = new Font(rchImg.SelectionFont, style);
        }
        private void cmdUnderline_Click(object sender, EventArgs e)
        {
            if (rchImg.SelectionFont == null)
            {
                return;
            }

            FontStyle style = rchImg.SelectionFont.Style;

            if (rchImg.SelectionFont.Underline)
            {
                style &= ~FontStyle.Underline;
            }
            else
            {
                style |= FontStyle.Underline;
            }
            rchImg.SelectionFont = new Font(rchImg.SelectionFont, style);
        }
        private void lstColors_DropDownItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            KnownColor selectedColor;
            selectedColor = (KnownColor)System.Enum.Parse(typeof(KnownColor), e.ClickedItem.Text);

            rchImg.SelectionColor = Color.FromKnownColor(selectedColor);
        }
        private void lstFonts_DropDownItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            if (rchImg.SelectionFont == null)
            {
                rchImg.SelectionFont = new Font(e.ClickedItem.Text, rchImg.Font.Size);
            }
            rchImg.SelectionFont = new Font(e.ClickedItem.Text, rchImg.SelectionFont.Size);
        }
        private void lstFontSize_DropDownItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            if (rchImg.SelectionFont == null)
            {
                return;
            }
            rchImg.SelectionFont = new Font(rchImg.SelectionFont.FontFamily,
                Convert.ToInt32(e.ClickedItem.Text),
                rchImg.SelectionFont.Style);
        }
        private void lstZoom_DropDownItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            float zoomPercent = Convert.ToSingle(e.ClickedItem.Text.Trim('%'));
            rchImg.ZoomFactor = zoomPercent / 100;
        }

        private void btnLogo_Click(object sender, EventArgs e)
        {
            Image img = Image.FromFile(_clsDef.PATHLOGHI + "Logo.jpg");
            Clipboard.SetImage(img);

            rchImg.SelectionStart = 0;
            rchImg.Paste();

            Clipboard.Clear();
        }

        private void btnAllergeni_Click(object sender, EventArgs e)
        {
            Allergeni2Upper();
        }

        private void Allergeni2Upper()
        {
            string sFil = "C:\\ApProject\\Temp\\Ingredienti\\Allergeni.txt";

            if (!File.Exists(sFil))
                MessageBox.Show(sFil + " non presente!", "FILE ALLEGERNI", MessageBoxButtons.OK, MessageBoxIcon.Error);
            else
            {
                string sRch = rchImg.Text;

                try
                {   // Open the text file using a stream reader.
                    using (StreamReader sr = new StreamReader(sFil))
                    {
                        // Read the stream to a string, and write the string to the console.
                        String line = sr.ReadToEnd();
                        Console.WriteLine(line);

                        string[] a = line.Split(',');
                        foreach (string word in a)
                        {
                            //Regex.Replace(word, @"[^\w\.@-]", "",
                            //        RegexOptions.None, TimeSpan.FromSeconds(1.5));

                            string s = Regex.Replace(word, @"[^\w\.@-]", "").ToLower();

                            Console.WriteLine(word);
                            if(s.Length > 2)
                                sRch = sRch.Replace(s, s.ToUpper());
                        }

                        rchImg.Text = sRch;

                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine("The file could not be read:");
                    Console.WriteLine(e.Message);
                }
            }
        }

        private void btnImport_Click(object sender, EventArgs e)
        {

            frmSeekArt f = new frmSeekArt();
            f.ShowDialog();
            if (f._tabArt != null && f._tabArt.Rows.Count > 0)
            {
                string sArt = (string)f._tabArt.Rows[0]["tmp_art"];

                string sFil = _clsDef.PATHINGREDIENTI + "et01_" + sArt + ".rtf";

                if (File.Exists(sFil))
                {
                    RichTextBox rchTmp = new RichTextBox();
                    rchTmp.LoadFile(sFil);

                    //rchImg.LoadFile(sFil);
                    rchImg.AppendText(rchTmp.Text);
                    rchTmp.Dispose();
                }
                else
                    MessageBox.Show("Ingredienti non trovati per l'articolo " + sArt + "!", "IMPORT INGREDIENTI", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

            }

        }

    }
}
