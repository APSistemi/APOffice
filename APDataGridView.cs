using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using System.Text;
using System.Diagnostics;
using Zuby.ADGV;

namespace APOffice
{
    public class APDataGridView : AdvancedDataGridView
    {
        private ContextMenuStrip _ctxtMenu;

        public APDataGridView()
        {
            this.AllowUserToOrderColumns = true;
            this.DoubleBuffered = true;

            // UI Refinement: Use extra padding to ensure the filter icon NEVER overlaps text.
            // Restoring native painting to preserve all filtering/sorting features.
            this.EnableHeadersVisualStyles = false;
            this.ColumnHeadersDefaultCellStyle.Font = new Font("Arial Narrow", 8F, FontStyle.Regular);
            this.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(245, 245, 245);
            this.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.TopLeft;
            this.ColumnHeadersDefaultCellStyle.Padding = new Padding(0, 2, 0, 0); // User requested 0 left margin
            this.ColumnHeadersHeight = 36;

            InitializeCustomFeatures();
        }

        private void InitializeCustomFeatures()
        {
            _ctxtMenu = new ContextMenuStrip();

            var mnuCopy = new ToolStripMenuItem("Copia valore");
            mnuCopy.Click += MnuCopy_Click;

            _ctxtMenu.Items.Add(mnuCopy);
            _ctxtMenu.Items.Add(new ToolStripSeparator());

            var mnuExport = new ToolStripMenuItem("Esporta dati in Excel...");
            mnuExport.Image = SystemIcons.WinLogo.ToBitmap();
            mnuExport.Click += MnuExport_Click;

            _ctxtMenu.Items.Add(mnuExport);

            var mnuClearFilter = new ToolStripMenuItem("Azzera filtri e ordinamento");
            mnuClearFilter.Click += MnuClearFilter_Click;

            _ctxtMenu.Items.Add(mnuClearFilter);

            _ctxtMenu.Opening += _ctxtMenu_Opening;
            this.ContextMenuStrip = _ctxtMenu;

            // Wire up the filter/sort events to auto-apply to BindingSource if present
            this.FilterStringChanged += APDataGridView_FilterStringChanged;
            this.SortStringChanged += APDataGridView_SortStringChanged;
        }

        private void APDataGridView_SortStringChanged(object sender, EventArgs e)
        {
            if (this.DataSource is BindingSource bs)
            {
                bs.Sort = this.SortString;
            }
        }

        private void APDataGridView_FilterStringChanged(object sender, EventArgs e)
        {
            if (this.DataSource is BindingSource bs)
            {
                bs.Filter = this.FilterString;
            }
        }

        private void _ctxtMenu_Opening(object sender, CancelEventArgs e)
        {
            // Abilita "Copia valore" solo se c'è una cella selezionata valida
            _ctxtMenu.Items[0].Enabled = (this.CurrentCell != null && this.CurrentCell.Value != null && this.CurrentCell.Value != DBNull.Value);
        }

        private void MnuCopy_Click(object sender, EventArgs e)
        {
            if (this.CurrentCell != null && this.CurrentCell.Value != null && this.CurrentCell.Value != DBNull.Value)
            {
                string testo = this.CurrentCell.Value.ToString();
                // Retry fino a 5 volte: ExternalException 0x800401D0 = clipboard temporaneamente bloccata da un altro processo
                for (int retry = 0; retry < 5; retry++)
                {
                    try
                    {
                        Clipboard.SetText(testo);
                        break; // successo
                    }
                    catch (System.Runtime.InteropServices.ExternalException)
                    {
                        if (retry == 4) break; // fallisce silenziosamente dopo 5 tentativi
                        System.Threading.Thread.Sleep(50);
                    }
                }
            }
        }

        private void MnuClearFilter_Click(object sender, EventArgs e)
        {
            this.CleanFilter();
            this.CleanSort();
            if (this.DataSource is BindingSource bs)
            {
                bs.Filter = null;
                bs.Sort = null;
            }
        }

        private void MnuExport_Click(object sender, EventArgs e)
        {
            try
            {
                ExportToExcel();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Errore durante l'esportazione: " + ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void ExportToExcel()
        {
            if (this.Rows.Count == 0)
            {
                MessageBox.Show("Nessun dato da esportare.", "Avviso", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string path = @"C:\APproject\XLS";
            if (!Directory.Exists(path)) Directory.CreateDirectory(path);

            string fileName = Path.Combine(path, "Export_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".xls");

            StringBuilder sb = new StringBuilder();

            // SpreadsheetML Header
            sb.AppendLine("<?xml version=\"1.0\"?>");
            sb.AppendLine("<?mso-application progid=\"Excel.Sheet\"?>");
            sb.AppendLine("<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\"");
            sb.AppendLine(" xmlns:o=\"urn:schemas-microsoft-com:office:office\"");
            sb.AppendLine(" xmlns:x=\"urn:schemas-microsoft-com:office:excel\"");
            sb.AppendLine(" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\"");
            sb.AppendLine(" xmlns:html=\"http://www.w3.org/TR/REC-html40\">");

            sb.AppendLine(" <Styles>");
            sb.AppendLine("  <Style ss:ID=\"Default\" ss:Name=\"Normal\">");
            sb.AppendLine("   <Alignment ss:Vertical=\"Bottom\"/>");
            sb.AppendLine("   <Borders/>");
            sb.AppendLine("   <Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Size=\"11\" ss:Color=\"#000000\"/>");
            sb.AppendLine("   <Interior/>");
            sb.AppendLine("   <NumberFormat/>");
            sb.AppendLine("   <Protection/>");
            sb.AppendLine("  </Style>");
            sb.AppendLine("  <Style ss:ID=\"Header\">");
            sb.AppendLine("   <Font ss:FontName=\"Calibri\" x:Family=\"Swiss\" ss:Size=\"11\" ss:Color=\"#FFFFFF\" ss:Bold=\"1\"/>");
            sb.AppendLine("   <Interior ss:Color=\"#4F81BD\" ss:Pattern=\"Solid\"/>");
            sb.AppendLine("  </Style>");
            sb.AppendLine(" </Styles>");

            sb.AppendLine(" <Worksheet ss:Name=\"DatiEsportati\">");
            sb.AppendLine("  <Table>");

            // Columns definition
            foreach (DataGridViewColumn col in this.Columns)
            {
                if (col.Visible)
                {
                    double width = col.Width * 0.75; // rough conversion px to points
                    sb.AppendLine(string.Format("   <Column ss:AutoFitWidth=\"0\" ss:Width=\"{0}\"/>", width));
                }
            }

            // Header Row
            sb.AppendLine("   <Row ss:AutoFitHeight=\"0\" ss:Height=\"20\">");
            foreach (DataGridViewColumn col in this.Columns)
            {
                if (col.Visible)
                {
                    sb.AppendLine("    <Cell ss:StyleID=\"Header\"><Data ss:Type=\"String\">" + EscapeXml(col.HeaderText) + "</Data></Cell>");
                }
            }
            sb.AppendLine("   </Row>");

            // Data Rows
            foreach (DataGridViewRow row in this.Rows)
            {
                if (row.IsNewRow) continue;

                sb.AppendLine("   <Row>");
                foreach (DataGridViewColumn col in this.Columns)
                {
                    if (col.Visible)
                    {
                        object val = row.Cells[col.Index].Value;
                        string cellType = "String";
                        string cellValue = "";

                        if (val != null && val != DBNull.Value)
                        {
                            if (val is double || val is decimal || val is float || val is int || val is long)
                            {
                                cellType = "Number";
                                cellValue = Convert.ToString(val, System.Globalization.CultureInfo.InvariantCulture);
                            }
                            else if (val is DateTime dt)
                            {
                                cellType = "DateTime";
                                cellValue = dt.ToString("yyyy-MM-ddTHH:mm:ss.fff");
                            }
                            else
                            {
                                cellValue = EscapeXml(val.ToString());
                            }
                        }

                        sb.AppendLine(string.Format("    <Cell><Data ss:Type=\"{0}\">{1}</Data></Cell>", cellType, cellValue));
                    }
                }
                sb.AppendLine("   </Row>");
            }

            sb.AppendLine("  </Table>");
            sb.AppendLine(" </Worksheet>");
            sb.AppendLine("</Workbook>");

            File.WriteAllText(fileName, sb.ToString(), Encoding.UTF8);

            // Open the file
            Process.Start(new ProcessStartInfo(fileName) { UseShellExecute = true });
        }

        protected override void OnDataError(bool displayErrorDialogIfNoHandler, DataGridViewDataErrorEventArgs e)
        {
            // Silently ignore standard data errors to prevent crash dialogs
            e.Cancel = true;
            e.ThrowException = false;
        }

        protected override void OnCellClick(DataGridViewCellEventArgs e)
        {
            if (this.RowCount == 0 && e.RowIndex == -1)
            {
                // Zuby ADGV can crash if header filter is clicked on an empty datagrid
                return;
            }
            base.OnCellClick(e);
        }

        private string EscapeXml(string str)
        {
            if (string.IsNullOrEmpty(str)) return "";
            return str.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("\"", "&quot;").Replace("'", "&apos;");
        }
    }
}
