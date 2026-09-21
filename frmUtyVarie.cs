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
    public partial class frmUtyVarie : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        private string _strConSql = "";
        private string _strConSqlLog = "";
        private string _strConSqlSta = "";

        public frmUtyVarie()
        {
            InitializeComponent();
            this.Shown += (s, ev) => frmWait.CloseWait();
            new clsGesGraph().SetGraph(this, 0);
            SetIcons();
        }

        private void frmUtyVarie_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");
            _strConSqlLog = _clsFun.ConSql("4");
            _strConSqlSta = _clsFun.ConSql("3");
            UpdateDashboard();
        }

        private void UpdateDashboard()
        {
            UpdateStat(_strConSql, lblStatOff, "APOffice");
            UpdateStat(_strConSqlLog, lblStatLog, "APLog");
            UpdateStat(_strConSqlSta, lblStatSta, "APStat");
        }

        private void UpdateStat(string sCon, Label lbl, string sDb)
        {
            try
            {
                // Caricamento istantaneo della dimensione del database
                string sSqlSize = "SELECT SUM(size) * 8 / 1024 FROM sys.master_files WHERE database_id = DB_ID()";
                DataTable tSize = _clsFun.FillTabSql("Size", sSqlSize, true, sCon);
                string sSize = (tSize != null && tSize.Rows.Count > 0) ? Convert.ToString(tSize.Rows[0][0]) : "-";

                lbl.Text = "Dimensione: " + sSize + " MB";

                // Calcolo frammentazione indici in background per non bloccare mai l'interfaccia utente (0 ms latenza UI)
                System.Threading.Tasks.Task.Run(() =>
                {
                    try
                    {
                        string sSqlFrag = "SELECT AVG(avg_fragmentation_in_percent) FROM sys.dm_db_index_physical_stats(DB_ID(), NULL, NULL, NULL, 'LIMITED') WHERE index_id > 0";
                        DataTable tFrag = _clsFun.FillTabSql("Frag", sSqlFrag, true, sCon);
                        string sFrag = "-";
                        if (tFrag != null && tFrag.Rows.Count > 0 && !DBNull.Value.Equals(tFrag.Rows[0][0]))
                            sFrag = Math.Round(Convert.ToDecimal(tFrag.Rows[0][0]), 2).ToString();

                        if (lbl.IsHandleCreated && !lbl.IsDisposed)
                        {
                            lbl.BeginInvoke(new Action(() =>
                            {
                                lbl.Text = "Dimensione: " + sSize + " MB | Frag.: " + sFrag + " %";
                            }));
                        }
                    }
                    catch { }
                });
            }
            catch (Exception ex)
            {
                lbl.Text = "Errore nel caricamento statistiche";
                _clsFun.ErrorLog("UpdateStat " + sDb, ex.Message);
            }
        }

        private void SetIcons()
        {
            button1.Image = GetIcon("📁", Color.DimGray);
            btnBkpOff.Image = GetIcon("💾", Color.SteelBlue);
            btnOptOff.Image = GetIcon("🚀", Color.ForestGreen);
            btnShrOff.Image = GetIcon("📉", Color.Sienna);
            btnBkpLog.Image = GetIcon("💾", Color.SteelBlue);
            btnOptLog.Image = GetIcon("🚀", Color.ForestGreen);
            btnShrLog.Image = GetIcon("📉", Color.Sienna);
            btnBkpSta.Image = GetIcon("💾", Color.SteelBlue);
            btnOptSta.Image = GetIcon("🚀", Color.ForestGreen);
            btnShrSta.Image = GetIcon("📉", Color.Sienna);
            btnDelNoEan.Image = GetIcon("🏷️", Color.Crimson);
            btnDelNoPrices.Image = GetIcon("💰", Color.DarkOrange);
            btnAlignApShop.Image = GetIcon("🔄", Color.RoyalBlue);

            foreach (Control c in this.Controls)
            {
                if (c is Button btn)
                {
                    btn.ImageAlign = ContentAlignment.MiddleLeft;
                    btn.TextImageRelation = TextImageRelation.ImageBeforeText;
                    btn.Padding = new Padding(10, 0, 0, 0);
                    btn.TextAlign = ContentAlignment.MiddleCenter;
                }
            }
        }

        private Image GetIcon(string unicodeChar, Color fallbackColor)
        {
            Bitmap bmp = new Bitmap(24, 24);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;
                using (Font f = new Font("Segoe UI Emoji", 12, FontStyle.Regular))
                using (SolidBrush b = new SolidBrush(fallbackColor))
                {
                    StringFormat fmt = new StringFormat();
                    fmt.Alignment = StringAlignment.Center;
                    fmt.LineAlignment = StringAlignment.Center;
                    g.DrawString(unicodeChar, f, b, new RectangleF(0, 0, 24, 24), fmt);
                }
            }
            return bmp;
        }

        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            CtrlDirectory();
        }

        private void btnDb_Click(object sender, EventArgs e)
        {
            Button btn = (Button)sender;
            string sDb = "";
            string sCon = "";
            string sCmd = "";

            if (btn.Name.Contains("Off")) { sDb = "APOffice"; sCon = _strConSql; }
            else if (btn.Name.Contains("Log")) { sDb = "APLog"; sCon = _strConSqlLog; }
            else if (btn.Name.Contains("Sta")) { sDb = "APStat"; sCon = _strConSqlSta; }

            if (btn.Name.Contains("Bkp")) { sCmd = "BACKUP"; }
            else if (btn.Name.Contains("Opt")) { sCmd = "OPTIMIZE"; }
            else if (btn.Name.Contains("Shr")) { sCmd = "SHRINK"; }

            if (sCmd == "BACKUP") BackupDb(sCon, sDb);
            else if (sCmd == "OPTIMIZE") OptimizeDb(sCon, sDb);
            else if (sCmd == "SHRINK") ShrinkDb(sCon, sDb);
        }

        private void BackupDb(string sCon, string sDb)
        {
            FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.Description = "Seleziona la cartella per il backup di " + sDb;

            if (!Directory.Exists(@"C:\APproject\Temp")) Directory.CreateDirectory(@"C:\APproject\Temp");
            fbd.SelectedPath = @"C:\APproject\Temp";

            if (fbd.ShowDialog() == DialogResult.OK)
            {
                string sPath = fbd.SelectedPath;
                if (!sPath.EndsWith(@"\")) sPath += @"\";
                string sFile = sDb + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".bak";
                string sFullName = sPath + sFile;

                string sSql = "BACKUP DATABASE [" + sDb + "] TO DISK = '" + sFullName + "'";
                try
                {
                    this.Cursor = Cursors.WaitCursor;
                    _clsFun.SqlWrite(sSql, sCon);
                    UpdateDashboard();
                    MessageBox.Show("Backup completato con successo!\n\nFile: " + sFile, "BACKUP DATABASE", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Errore durante il backup:\n" + ex.Message, "ERRORE BACKUP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    this.Cursor = Cursors.Default;
                }
            }
        }

        private void OptimizeDb(string sCon, string sDb)
        {
            string sSql = @"
SET NOCOUNT ON;
DECLARE @TableName NVARCHAR(255);
DECLARE @IndexName NVARCHAR(255);
DECLARE @Fragmentation FLOAT;
DECLARE @SQL NVARCHAR(MAX);

DECLARE index_cursor CURSOR FOR
SELECT 
    t.name AS TableName,
    i.name AS IndexName,
    stats.avg_fragmentation_in_percent
FROM sys.dm_db_index_physical_stats(DB_ID(), NULL, NULL, NULL, 'LIMITED') AS stats
INNER JOIN sys.tables AS t ON stats.object_id = t.object_id
INNER JOIN sys.indexes AS i ON stats.object_id = i.object_id AND stats.index_id = i.index_id
WHERE stats.avg_fragmentation_in_percent > 10.0 
  AND i.name IS NOT NULL;

OPEN index_cursor;
FETCH NEXT FROM index_cursor INTO @TableName, @IndexName, @Fragmentation;

WHILE @@FETCH_STATUS = 0
BEGIN
    IF @Fragmentation > 30.0
        SET @SQL = 'ALTER INDEX [' + @IndexName + '] ON [' + @TableName + '] REBUILD';
    ELSE
        SET @SQL = 'ALTER INDEX [' + @IndexName + '] ON [' + @TableName + '] REORGANIZE';
    
    EXEC sp_executesql @SQL;
    FETCH NEXT FROM index_cursor INTO @TableName, @IndexName, @Fragmentation;
END;

CLOSE index_cursor;
DEALLOCATE index_cursor;

EXEC sp_msforeachtable 'UPDATE STATISTICS ?';
";
            try
            {
                this.Cursor = Cursors.WaitCursor;
                _clsFun.SqlWrite(sSql, sCon);
                UpdateDashboard();
                MessageBox.Show("Ottimizzazione completata per " + sDb + "!", "OTTIMIZZAZIONE DB", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Errore durante l'ottimizzazione:\n" + ex.Message, "ERRORE OTTIMIZZAZIONE", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        private void ShrinkDb(string sCon, string sDb)
        {
            string sSql = "DBCC SHRINKDATABASE ([" + sDb + "])";
            try
            {
                this.Cursor = Cursors.WaitCursor;
                _clsFun.SqlWrite(sSql, sCon);
                UpdateDashboard();
                MessageBox.Show("Compattazione completata per " + sDb + "!", "SHRINK DATABASE", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Errore durante lo shrink:\n" + ex.Message, "ERRORE SHRINK", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        private void CtrlDirectory()
        {
            string sPath = Path.GetDirectoryName(_clsDef.FILEDIR) + "\\";

            if (!File.Exists(_clsDef.FILEDIR))
                MessageBox.Show(_clsDef.FILEDIR + " non presente!");
            else
            {
                string sRig = "";

                using (StreamReader sr = new StreamReader(_clsDef.FILEDIR))
                {
                    while ((sRig = sr.ReadLine()) != null)
                    {

                        string s = sPath + sRig;

                        if (!Directory.Exists(s))
                            System.IO.Directory.CreateDirectory(s);
                    }
                    sr.Close();
                    sr.Dispose();
                    MessageBox.Show("Fatto!");
                }
            }
        }

        private int ExecuteSqlStatements(string[] sqlStatements, string sCon)
        {
            int totalAffected = 0;
            using (System.Data.SqlClient.SqlConnection cn = new System.Data.SqlClient.SqlConnection(sCon))
            {
                cn.Open();
                using (System.Data.SqlClient.SqlTransaction tr = cn.BeginTransaction())
                {
                    try
                    {
                        foreach (string sql in sqlStatements)
                        {
                            if (string.IsNullOrWhiteSpace(sql)) continue;
                            using (System.Data.SqlClient.SqlCommand cm = new System.Data.SqlClient.SqlCommand(sql, cn, tr))
                            {
                                cm.CommandTimeout = 300;
                                totalAffected += cm.ExecuteNonQuery();
                            }
                        }
                        tr.Commit();
                    }
                    catch (Exception ex)
                    {
                        tr.Rollback();
                        throw ex;
                    }
                }
            }
            return totalAffected;
        }

        private void btnDelNoEan_Click(object sender, EventArgs e)
        {
            try
            {
                string sFilter = "art_cod NOT IN (SELECT DISTINCT ean_art FROM AnaBarcode WHERE ean_art IS NOT NULL AND ean_art <> '' AND ean_ean IS NOT NULL AND ean_ean <> '' AND (ean_ann = 0 OR ean_ann IS NULL))";
                string sSqlCount = "SELECT COUNT(*) FROM AnaArticoli WHERE " + sFilter;
                DataTable tCount = _clsFun.FillTabSql("AnaArticoli", sSqlCount, false, _strConSql);
                int nCount = (tCount != null && tCount.Rows.Count > 0 && !DBNull.Value.Equals(tCount.Rows[0][0])) ? Convert.ToInt32(tCount.Rows[0][0]) : 0;

                if (nCount == 0)
                {
                    MessageBox.Show("Nessun articolo sprovvisto di barcode (EAN) trovato.", "PULIZIA ARTICOLI", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                DialogResult res = MessageBox.Show(
                    "Trovati " + nCount.ToString("N0") + " articoli sprovvisti di barcode (EAN).\n\n" +
                    "ATTENZIONE: Verranno eliminati definitivamente gli articoli e tutti i riferimenti collegati (barcode, listini vendita, costi, ingredienti, variazioni).\n\n" +
                    "Confermi l'eliminazione?",
                    "CONFERMA ELIMINAZIONE ARTICOLI SENZA BARCODE",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);

                if (res == DialogResult.Yes)
                {
                    this.Cursor = Cursors.WaitCursor;

                    string[] sqlStatements = new string[]
                    {
                        "IF OBJECT_ID('AnaBarcode', 'U') IS NOT NULL DELETE FROM AnaBarcode WHERE ean_art IN (SELECT art_cod FROM AnaArticoli WHERE " + sFilter + ")",
                        "IF OBJECT_ID('GesLisVendita', 'U') IS NOT NULL DELETE FROM GesLisVendita WHERE liv_art IN (SELECT art_cod FROM AnaArticoli WHERE " + sFilter + ")",
                        "IF OBJECT_ID('GesLisAcquisto', 'U') IS NOT NULL DELETE FROM GesLisAcquisto WHERE lia_art IN (SELECT art_cod FROM AnaArticoli WHERE " + sFilter + ")",
                        "IF OBJECT_ID('AnaArtIngredienti', 'U') IS NOT NULL DELETE FROM AnaArtIngredienti WHERE ing_art IN (SELECT art_cod FROM AnaArticoli WHERE " + sFilter + ")",
                        "IF OBJECT_ID('GesVariazioni', 'U') IS NOT NULL DELETE FROM GesVariazioni WHERE var_art IN (SELECT art_cod FROM AnaArticoli WHERE " + sFilter + ")",
                        "IF OBJECT_ID('AnaVariazioni', 'U') IS NOT NULL DELETE FROM AnaVariazioni WHERE var_art IN (SELECT art_cod FROM AnaArticoli WHERE " + sFilter + ")",
                        "IF OBJECT_ID('AnaArtGiacenza', 'U') IS NOT NULL DELETE FROM AnaArtGiacenza WHERE gia_art IN (SELECT art_cod FROM AnaArticoli WHERE " + sFilter + ")",
                        "IF OBJECT_ID('AnaArticoli', 'U') IS NOT NULL DELETE FROM AnaArticoli WHERE " + sFilter
                    };

                    ExecuteSqlStatements(sqlStatements, _strConSql);

                    MessageBox.Show("Operazione completata con successo!\n\nEliminati " + nCount.ToString("N0") + " articoli senza barcode e relative dipendenze.", "PULIZIA COMPLETATA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    UpdateDashboard();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Errore durante la cancellazione degli articoli:\n" + ex.Message, "ERRORE PULIZIA", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _clsFun.ErrorLog("btnDelNoEan_Click", ex.Message);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        private void btnDelNoPrices_Click(object sender, EventArgs e)
        {
            try
            {
                string sFilter = "art_cod NOT IN (SELECT DISTINCT liv_art FROM GesLisVendita WHERE liv_art IS NOT NULL AND liv_art <> '' AND (liv_ann = 0 OR liv_ann IS NULL)) AND art_cod NOT IN (SELECT DISTINCT lia_art FROM GesLisAcquisto WHERE lia_art IS NOT NULL AND lia_art <> '' AND (lia_ann = 0 OR lia_ann IS NULL))";
                string sSqlCount = "SELECT COUNT(*) FROM AnaArticoli WHERE " + sFilter;
                DataTable tCount = _clsFun.FillTabSql("AnaArticoli", sSqlCount, false, _strConSql);
                int nCount = (tCount != null && tCount.Rows.Count > 0 && !DBNull.Value.Equals(tCount.Rows[0][0])) ? Convert.ToInt32(tCount.Rows[0][0]) : 0;

                if (nCount == 0)
                {
                    MessageBox.Show("Nessun articolo privo di listini di vendita e costi collegati trovato.", "PULIZIA ARTICOLI", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                DialogResult res = MessageBox.Show(
                    "Trovati " + nCount.ToString("N0") + " articoli privi di listini di vendita e costi collegati.\n\n" +
                    "ATTENZIONE: Verranno eliminati definitivamente gli articoli e tutti i riferimenti collegati (barcode, ingredienti, variazioni).\n\n" +
                    "Confermi l'eliminazione?",
                    "CONFERMA ELIMINAZIONE ARTICOLI SENZA PREZZI/COSTI",
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);

                if (res == DialogResult.Yes)
                {
                    this.Cursor = Cursors.WaitCursor;

                    string[] sqlStatements = new string[]
                    {
                        "IF OBJECT_ID('AnaBarcode', 'U') IS NOT NULL DELETE FROM AnaBarcode WHERE ean_art IN (SELECT art_cod FROM AnaArticoli WHERE " + sFilter + ")",
                        "IF OBJECT_ID('GesLisVendita', 'U') IS NOT NULL DELETE FROM GesLisVendita WHERE liv_art IN (SELECT art_cod FROM AnaArticoli WHERE " + sFilter + ")",
                        "IF OBJECT_ID('GesLisAcquisto', 'U') IS NOT NULL DELETE FROM GesLisAcquisto WHERE lia_art IN (SELECT art_cod FROM AnaArticoli WHERE " + sFilter + ")",
                        "IF OBJECT_ID('AnaArtIngredienti', 'U') IS NOT NULL DELETE FROM AnaArtIngredienti WHERE ing_art IN (SELECT art_cod FROM AnaArticoli WHERE " + sFilter + ")",
                        "IF OBJECT_ID('GesVariazioni', 'U') IS NOT NULL DELETE FROM GesVariazioni WHERE var_art IN (SELECT art_cod FROM AnaArticoli WHERE " + sFilter + ")",
                        "IF OBJECT_ID('AnaVariazioni', 'U') IS NOT NULL DELETE FROM AnaVariazioni WHERE var_art IN (SELECT art_cod FROM AnaArticoli WHERE " + sFilter + ")",
                        "IF OBJECT_ID('AnaArtGiacenza', 'U') IS NOT NULL DELETE FROM AnaArtGiacenza WHERE gia_art IN (SELECT art_cod FROM AnaArticoli WHERE " + sFilter + ")",
                        "IF OBJECT_ID('AnaArticoli', 'U') IS NOT NULL DELETE FROM AnaArticoli WHERE " + sFilter
                    };

                    ExecuteSqlStatements(sqlStatements, _strConSql);

                    MessageBox.Show("Operazione completata con successo!\n\nEliminati " + nCount.ToString("N0") + " articoli privi di listini/costi e relative dipendenze.", "PULIZIA COMPLETATA", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    UpdateDashboard();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Errore durante la cancellazione degli articoli:\n" + ex.Message, "ERRORE PULIZIA", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _clsFun.ErrorLog("btnDelNoPrices_Click", ex.Message);
            }
            finally
            {
                this.Cursor = Cursors.Default;
            }
        }

        private void btnAlignApShop_Click(object sender, EventArgs e)
        {
            try
            {
                string sDefaultPath = @"C:\ApProject\ApShop\DataBase\DataBase.mdb";

                frmUtyInput1 fInput = new frmUtyInput1();
                fInput._strDes = "Percorso Database APShop (MDB):";
                fInput._strTxt = sDefaultPath;
                fInput.ShowDialog(this);

                string sPath = fInput._strTxt.Trim();
                if (string.IsNullOrEmpty(sPath))
                    return;

                frmUtyAlignApShop frm = new frmUtyAlignApShop(sPath);
                if (frm.ShowDialog(this) == DialogResult.OK)
                {
                    UpdateDashboard();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Errore durante l'apertura della maschera di allineamento APShop:\n" + ex.Message, "ERRORE", MessageBoxButtons.OK, MessageBoxIcon.Error);
                _clsFun.ErrorLog("btnAlignApShop_Click", ex.Message);
            }
        }
    }
}
