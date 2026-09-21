using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Diagnostics;
using System.Linq;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using BarcodeLib;

namespace APOffice
{
    public class clsGenPdfEtiDB
    {
        private clsFuncs _clsFun = new clsFuncs();
        private clsDefine _clsDef = new clsDefine();
        private string _strConSql = "";

        public clsGenPdfEtiDB()
        {
            _strConSql = _clsFun.ConSql("");
        }

        public static void SanitizeUnitPrices(DataTable tabEti)
        {
            if (tabEti == null) return;
            foreach (DataRow row in tabEti.Rows)
            {
                try
                {
                    if (row.Table.Columns.Contains("eti_prv") && row.Table.Columns.Contains("eti_pne"))
                    {
                        decimal prv = (!DBNull.Value.Equals(row["eti_prv"])) ? Convert.ToDecimal(row["eti_prv"]) : 0m;
                        decimal pne = (!DBNull.Value.Equals(row["eti_pne"])) ? Convert.ToDecimal(row["eti_pne"]) : 0m;
                        decimal tgv = (row.Table.Columns.Contains("eti_tgv") && !DBNull.Value.Equals(row["eti_tgv"])) ? Convert.ToDecimal(row["eti_tgv"]) : 1m;

                        if (pne > 500m)
                        {
                            row["eti_pne"] = 0m;
                        }
                        else if (prv > 0m && pne > 0m && tgv > 0m)
                        {
                            decimal d = prv / (pne / tgv);
                            if (d > 500m)
                            {
                                row["eti_pne"] = 0m;
                            }
                        }
                    }
                    if (row.Table.Columns.Contains("pos_prv") && row.Table.Columns.Contains("pos_pne"))
                    {
                        decimal pne = (!DBNull.Value.Equals(row["pos_pne"])) ? Convert.ToDecimal(row["pos_pne"]) : 0m;
                        if (pne > 500m)
                        {
                            row["pos_pne"] = 0m;
                        }
                    }
                }
                catch { }
            }
        }

        public void PrnPdfEtiDaDB(DataTable tabEti, string strCodFormat, int skipLabels = 0)
        {
            SanitizeUnitPrices(tabEti);
            string s = "SELECT * FROM TabEtiFormatiLayout WHERE eti_cod='" + strCodFormat + "'";
            DataTable tLay = _clsFun.FillTabSql("TabEtiFormatiLayout", s, false, _strConSql);
            if (tLay.Rows.Count == 0) return;

            DataRow rowLay = tLay.Rows[0];

            int mmWidth = 100, mmHeight = 100;
            string etiDim = rowLay["eti_dim"].ToString();
            if (etiDim.Contains(","))
            {
                string[] dims = etiDim.Split(',');
                int.TryParse(dims[0], out mmWidth);
                int.TryParse(dims[1], out mmHeight);
            }

            int iCols = 1, iRows = 1;
            double topMargin = 0, leftMargin = 0, xStep = 0, yStep = 0;
            bool isLandscape = false;
            string etiPag = "";
            if (tLay.Columns.Contains("eti_pag")) etiPag = rowLay["eti_pag"].ToString();

            if (!string.IsNullOrEmpty(etiPag))
            {
                string[] pags = etiPag.Split(',');
                if (pags.Length >= 6)
                {
                    int.TryParse(pags[0], out iCols);
                    int.TryParse(pags[1], out iRows);
                    double.TryParse(pags[2], out topMargin);
                    double.TryParse(pags[3], out leftMargin);
                    double.TryParse(pags[4], out xStep);
                    double.TryParse(pags[5], out yStep);
                }
                if (pags.Length >= 7)
                    isLandscape = (pags[6].Trim().ToUpper() == "H");
            }
            if (iCols <= 0) iCols = 1;
            if (iRows <= 0) iRows = 1;

            // Convert mm to points (1 point = 1/72 inch, 1 inch = 25.4 mm)
            double ptWidth = mmWidth * 72.0 / 25.4;
            double ptHeight = mmHeight * 72.0 / 25.4;
            double ptTopMargin = topMargin * 72.0 / 25.4;
            double ptLeftMargin = leftMargin * 72.0 / 25.4;
            double ptXStep = xStep * 72.0 / 25.4;
            double ptYStep = yStep * 72.0 / 25.4;

            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Etichette Personalizzate";

            Barcode _ean13 = new BarcodeLib.Barcode();
            _ean13.Alignment = BarcodeLib.AlignmentPositions.CENTER;
            _ean13.IncludeLabel = true;
            _ean13.LabelPosition = BarcodeLib.LabelPositions.BOTTOMCENTER;

            string dirIni = Path.GetDirectoryName(_clsDef.FILEINI); // E.g., C:\ApProject\APOffice\DataBase
            string dirPdf = Path.GetFullPath(Path.Combine(dirIni, @"..\..\Pdf"));
            if (!Directory.Exists(dirPdf))
            {
                try { Directory.CreateDirectory(dirPdf); } catch { dirPdf = dirIni; } // Fallback se non ci sono i permessi
            }
            string filePdf = Path.Combine(dirPdf, "Etichette_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf");

            int curCol = 0, curRow = 0;
            PdfPage page = null;
            XGraphics gfx = null;

            // Avanzamento a vuoto iniziale (Salto Etichette per iniziare da "Riga inizio")
            for (int i = 0; i < skipLabels; i++)
            {
                if (page == null)
                {
                    page = pd.AddPage();
                    page.Size = PdfSharp.PageSize.A4;
                    if (isLandscape) page.Orientation = PdfSharp.PageOrientation.Landscape;
                    else page.Orientation = PdfSharp.PageOrientation.Portrait;
                    gfx = XGraphics.FromPdfPage(page);
                }

                curCol++;
                if (curCol >= iCols)
                {
                    curCol = 0;
                    curRow++;
                    if (curRow >= iRows)
                    {
                        page = pd.AddPage();
                        page.Size = PdfSharp.PageSize.A4;
                        if (isLandscape) page.Orientation = PdfSharp.PageOrientation.Landscape;
                        else page.Orientation = PdfSharp.PageOrientation.Portrait;
                        gfx = XGraphics.FromPdfPage(page);
                        curRow = 0;
                    }
                }
            }

            foreach (DataRow rowDati in tabEti.Rows)
            {
                if (rowDati["eti_inv"].ToString() == "1") continue;

                int copie = 1;
                try { if (!DBNull.Value.Equals(rowDati["eti_qta"])) copie = Convert.ToInt32(rowDati["eti_qta"]); } catch { }
                if (copie <= 0) copie = 1;

                for (int c = 0; c < copie; c++)
                {
                    if (page == null || curCol >= iCols)
                    {
                        if (page != null)
                        {
                            curCol = 0;
                            curRow++;
                        }
                        if (page == null || curRow >= iRows)
                        {
                            page = pd.AddPage();
                            page.Size = PdfSharp.PageSize.A4;
                            if (isLandscape)
                                page.Orientation = PdfSharp.PageOrientation.Landscape;
                            else
                                page.Orientation = PdfSharp.PageOrientation.Portrait;
                            gfx = XGraphics.FromPdfPage(page);
                            curCol = 0;
                            curRow = 0;
                        }
                    }

                    double offsetX = ptLeftMargin + (curCol * ptXStep);
                    double offsetY = ptTopMargin + (curRow * ptYStep);

                    // Disegna le Forme di Sfondo (ZIndex = 0)
                    DisegnaForme(gfx, rowLay, 0, offsetX, offsetY);

                    // Disegna le Immagini di Sfondo (ZIndex = 0 = Dietro)
                    DisegnaElementoImmagine(gfx, rowLay["eti_s012"].ToString(), offsetX, offsetY, "", 0);
                    string imgArtBg = Convert.ToString(rowDati["eti_img"]);
                    DisegnaElementoImmagine(gfx, rowLay["eti_s017"].ToString(), offsetX, offsetY, !string.IsNullOrEmpty(imgArtBg) ? imgArtBg : "", 0);

                    // Disegna i Testi di Sfondo (ZIndex = 0 = Dietro)
                    DisegnaTuttiITesti(gfx, rowLay, rowDati, offsetX, offsetY, 0);

                    // Barcode
                    string cfgBarcode = rowLay["eti_s013"].ToString();
                    if (!string.IsNullOrEmpty(cfgBarcode) && cfgBarcode.StartsWith("S"))
                    {
                        string[] a = cfgBarcode.Split(';');
                        if (a.Length >= 3)
                        {
                            string[] pos = a[1].Split(',');
                            if (pos.Length >= 4)
                            {
                                double bY = Convert.ToDouble(pos[0]) * 72.0 / 25.4 + offsetY;
                                double bX = Convert.ToDouble(pos[1]) * 72.0 / 25.4 + offsetX;
                                double bW = Convert.ToDouble(pos[2]) * 72.0 / 25.4;
                                double bH = Convert.ToDouble(pos[3]) * 72.0 / 25.4;

                                string ean = Convert.ToString(rowDati["eti_ean"]);
                                if (ean.Length >= 8)
                                {
                                    try
                                    {
                                        // Applica font personalizzato per il Barcode
                                        string[] fntBarcode = a[2].Split('|');
                                        string bFName = fntBarcode.Length > 0 && fntBarcode[0] != "" ? fntBarcode[0] : "Arial";
                                        float bFSize = fntBarcode.Length > 1 && fntBarcode[1] != "" ? Convert.ToSingle(fntBarcode[1]) : 10f;
                                        System.Drawing.FontStyle bFStyle = System.Drawing.FontStyle.Regular;
                                        if (fntBarcode.Length > 2)
                                        {
                                            if (fntBarcode[2].ToUpper() == "B") bFStyle = System.Drawing.FontStyle.Bold;
                                            if (fntBarcode[2].ToUpper() == "I") bFStyle = System.Drawing.FontStyle.Italic;
                                        }
                                        // Il font va scalato * 2 perché l'immagine viene generata al doppio della risoluzione
                                        _ean13.LabelFont = new System.Drawing.Font(bFName, bFSize * 2f, bFStyle);

                                        using (System.Drawing.Image img = _ean13.Encode(ean.Length == 13 ? BarcodeLib.TYPE.EAN13 : BarcodeLib.TYPE.EAN8, ean, Color.Black, Color.White, (int)(bW * 2), (int)(bH * 2)))
                                        {
                                            XImage xImg = XImage.FromGdiPlusImage(img);
                                            gfx.DrawImage(xImg, bX, bY, bW, bH);
                                        }
                                    }
                                    catch { }
                                }
                            }
                        }
                    }

                    // Disegna le Immagini in Primo Piano (ZIndex = 1 = Sopra)
                    DisegnaElementoImmagine(gfx, rowLay["eti_s012"].ToString(), offsetX, offsetY, "", 1);
                    string imgArtFg = Convert.ToString(rowDati["eti_img"]);
                    DisegnaElementoImmagine(gfx, rowLay["eti_s017"].ToString(), offsetX, offsetY, !string.IsNullOrEmpty(imgArtFg) ? imgArtFg : "", 1);

                    // Disegna i Testi in Primo Piano (ZIndex = 1 = Sopra)
                    DisegnaTuttiITesti(gfx, rowLay, rowDati, offsetX, offsetY, 1);

                    // Disegna le Forme/Cornici in Primo Piano (ZIndex = 1)
                    DisegnaForme(gfx, rowLay, 1, offsetX, offsetY);

                    // Disegna i Testi "Sopra Tutto" (ZIndex = 2 = In primo piano assoluto)
                    DisegnaTuttiITesti(gfx, rowLay, rowDati, offsetX, offsetY, 2);

                    curCol++;
                }
            }

            if (pd.PageCount > 0)
            {
                pd.Save(filePdf);
                Process.Start(filePdf);
            }
        }

        private void DisegnaForme(XGraphics gfx, DataRow rowLay, int drawZIndex, double offsetX, double offsetY)
        {
            for (int i = 1; i <= 8; i++)
            {
                string colName = "eti_h00" + i;
                if (!rowLay.Table.Columns.Contains(colName)) continue;
                string raw = rowLay[colName].ToString();
                if (string.IsNullOrEmpty(raw) || !raw.StartsWith("S")) continue;

                string[] parts = raw.Split(';');

                // Leggi ZIndex: se manca il 4° segmento (record legacy), default = 1 (Primo Piano, sopra le immagini)
                string[] shp = new string[0];
                int zIndex = 1;
                if (parts.Length >= 4)
                {
                    shp = parts[3].Split('|');
                    if (shp.Length > 0) int.TryParse(shp[0], out zIndex);
                }
                if (zIndex != drawZIndex) continue;

                // Coordinate e Dimensioni (secondo segmento)
                string[] dim = parts[1].Split(',');
                if (dim.Length < 4) continue;

                double y = Convert.ToDouble(dim[0]) * 72.0 / 25.4 + offsetY;
                double x = Convert.ToDouble(dim[1]) * 72.0 / 25.4 + offsetX;
                double w = Convert.ToDouble(dim[2]) * 72.0 / 25.4;
                double h = Convert.ToDouble(dim[3]) * 72.0 / 25.4;

                int shapeColor = Color.Black.ToArgb();
                int shapeThick = 1;
                bool shapeFill = false;
                int shapeType = 1;

                if (shp.Length > 1) int.TryParse(shp[1], out shapeColor);
                if (shp.Length > 2) int.TryParse(shp[2], out shapeThick);
                shapeFill = (shp.Length > 3 && shp[3] == "F");
                if (shp.Length > 4) int.TryParse(shp[4], out shapeType);

                Color c = Color.FromArgb(shapeColor == 0 ? Color.Black.ToArgb() : shapeColor);
                XColor xc = XColor.FromArgb(255, c.R, c.G, c.B); // Ignoro alpha, forzando opacità piena

                if (shapeType == 2) // Linea Orizzontale
                {
                    XPen pen = new XPen(xc, shapeThick);
                    gfx.DrawLine(pen, x, y + h / 2, x + w, y + h / 2);
                }
                else if (shapeType == 3) // Linea Verticale
                {
                    XPen pen = new XPen(xc, shapeThick);
                    gfx.DrawLine(pen, x + w / 2, y, x + w / 2, y + h);
                }
                else // 1 = Rettangolo o default
                {
                    if (shapeFill)
                    {
                        XBrush br = new XSolidBrush(xc);
                        gfx.DrawRectangle(br, x, y, w, h);
                    }
                    else
                    {
                        XPen pen = new XPen(xc, shapeThick);
                        gfx.DrawRectangle(pen, x, y, w, h);
                    }
                }
            }
        }

        // Helper: legge ZIndex dal quarto segmento della stringa cfg
        private int GetCfgZIndex(string cfg)
        {
            if (string.IsNullOrEmpty(cfg)) return 1;
            string[] parts = cfg.Split(';');
            if (parts.Length < 4) return 1;
            string[] shp = parts[3].Split('|');
            int z = 1;
            if (shp.Length > 0) int.TryParse(shp[0], out z);
            return z;
        }

        private void DisegnaElementoImmagine(XGraphics gfx, string cfg, double offsetX = 0, double offsetY = 0, string overridePath = "", int zIndexFilter = -1)
        {
            if (string.IsNullOrEmpty(cfg) || !cfg.StartsWith("S")) return;

            // Filtro ZIndex: -1 = sempre, altrimenti deve corrispondere
            if (zIndexFilter >= 0 && GetCfgZIndex(cfg) != zIndexFilter) return;

            string[] a = cfg.Split(';');
            if (a.Length < 3) return;

            // Coordinate (secondo segmento)
            string[] pos = a[1].Split(',');
            if (pos.Length < 4) return;

            double y = Convert.ToDouble(pos[0]) * 72.0 / 25.4 + offsetY;
            double x = Convert.ToDouble(pos[1]) * 72.0 / 25.4 + offsetX;
            double w = Convert.ToDouble(pos[2]) * 72.0 / 25.4;
            double h = Convert.ToDouble(pos[3]) * 72.0 / 25.4;

            // Path immagine (terzo segmento, prima del primo '|') o override
            string path = overridePath;
            if (string.IsNullOrEmpty(path))
            {
                string[] imgData = a[2].Split('|');
                if (imgData.Length > 0) path = imgData[0];
            }

            if (!string.IsNullOrEmpty(path) && File.Exists(path))
            {
                try
                {
                    using (XImage image = XImage.FromFile(path))
                    {
                        gfx.DrawImage(image, x, y, w, h);
                    }
                }
                catch { }
            }
        }

        private void DisegnaElementoTesto(XGraphics gfx, string cfg, string testo, double offsetX = 0, double offsetY = 0, int zIndexFilter = -1)
        {
            if (string.IsNullOrEmpty(cfg) || !cfg.StartsWith("S")) return;
            if (string.IsNullOrEmpty(testo)) return;

            // Filtro ZIndex
            if (zIndexFilter >= 0 && GetCfgZIndex(cfg) != zIndexFilter) return;

            string[] a = cfg.Split(';');
            if (a.Length < 3) return;

            // Formato standard Y,X,W,H,Align (secondo segmento)
            string[] pos = a[1].Split(',');
            if (pos.Length < 4) return;

            double y = Convert.ToDouble(pos[0]) * 72.0 / 25.4 + offsetY;
            double x = Convert.ToDouble(pos[1]) * 72.0 / 25.4 + offsetX;
            double w = Convert.ToDouble(pos[2]) * 72.0 / 25.4;
            double h = Convert.ToDouble(pos[3]) * 72.0 / 25.4;

            XStringFormat format = new XStringFormat();
            format.Alignment = XStringAlignment.Near;
            format.LineAlignment = XLineAlignment.Near;

            if (pos.Length > 4)
            {
                if (pos[4] == "1") format.Alignment = XStringAlignment.Far;
                if (pos[4] == "2") format.Alignment = XStringAlignment.Center;
            }

            // Formato font: FontName|Size|Style (terzo segmento)
            string[] fnt = a[2].Split('|');
            string fName = fnt.Length > 0 && fnt[0] != "" ? fnt[0] : "Arial";
            double fSize = fnt.Length > 1 && fnt[1] != "" ? Convert.ToDouble(fnt[1]) : 10;
            XFontStyle fStyle = XFontStyle.Regular;
            if (fnt.Length > 2)
            {
                if (fnt[2].ToUpper() == "B") fStyle = XFontStyle.Bold;
                if (fnt[2].ToUpper() == "I") fStyle = XFontStyle.Italic;
            }

            XFont font = new XFont(fName, fSize, fStyle);
            XRect rect = new XRect(x, y, w, h);

            int fontColorArgb = Color.Black.ToArgb();
            if (a.Length > 3)
            {
                string[] shp = a[3].Split('|');
                if (shp.Length > 1) int.TryParse(shp[1], out fontColorArgb);
            }
            XSolidBrush fBrush = new XSolidBrush(XColor.FromArgb(fontColorArgb));

            gfx.DrawString(testo, font, fBrush, rect, format);
        }

        private void DisegnaTuttiITesti(XGraphics gfx, DataRow rowLay, DataRow rowDati, double offsetX, double offsetY, int zIndexFilter)
        {
            // [s003] Intestazione / Ragione Sociale
            string s003val = Convert.ToString(rowDati["eti_ard"]);
            DisegnaElementoTesto(gfx, rowLay["eti_s003"].ToString(), s003val, offsetX, offsetY, zIndexFilter);

            // [s004] Descrizione articolo
            DisegnaDescrizioneFit(gfx, rowLay["eti_s004"].ToString(), Convert.ToString(rowDati["eti_ard"]), offsetX, offsetY, zIndexFilter);

            // [s005] Info addizionali
            string s005val = "";
            if (!DBNull.Value.Equals(rowDati["eti_bor"])) s005val = Convert.ToString(rowDati["eti_bor"]);
            DisegnaElementoTesto(gfx, rowLay["eti_s005"].ToString(), s005val, offsetX, offsetY, zIndexFilter);

            // [s006] Codice Interno articolo
            DisegnaElementoTesto(gfx, rowLay["eti_s006"].ToString(), Convert.ToString(rowDati["eti_art"]), offsetX, offsetY, zIndexFilter);

            // [s007] Slogan / Note promo
            string s007val = "";
            if (!DBNull.Value.Equals(rowDati["eti_cam"])) s007val = Convert.ToString(rowDati["eti_cam"]);
            DisegnaElementoTesto(gfx, rowLay["eti_s007"].ToString(), s007val, offsetX, offsetY, zIndexFilter);

            // [s008] Prezzo al Kg / Netto
            string s008val = "";
            if (!DBNull.Value.Equals(rowDati["eti_pne"]) && Convert.ToDecimal(rowDati["eti_pne"]) > 0 && Convert.ToDecimal(rowDati["eti_pne"]) <= 500m)
                s008val = "€ " + Convert.ToDecimal(rowDati["eti_pne"]).ToString("0.00") + "/kg";
            DisegnaElementoTesto(gfx, rowLay["eti_s008"].ToString(), s008val, offsetX, offsetY, zIndexFilter);

            // [s010] Prezzo VECCHIO
            string s010val = "";
            decimal pve = 0, prv = 0;
            if (!DBNull.Value.Equals(rowDati["eti_pve"])) pve = Convert.ToDecimal(rowDati["eti_pve"]);
            if (!DBNull.Value.Equals(rowDati["eti_prv"])) prv = Convert.ToDecimal(rowDati["eti_prv"]);
            if (pve > 0 && pve != prv) s010val = "€ " + pve.ToString("0.00");
            DisegnaElementoTesto(gfx, rowLay["eti_s010"].ToString(), s010val, offsetX, offsetY, zIndexFilter);

            // [s009] Sconto / Risparmio in Euro
            string s009val = "";
            if (pve > 0 && pve > prv) s009val = "Sconto € " + (pve - prv).ToString("0.00");
            DisegnaElementoTesto(gfx, rowLay["eti_s009"].ToString(), s009val, offsetX, offsetY, zIndexFilter);

            // [s011] PREZZO FINALE di vendita
            if (prv > 0) DisegnaPrezzoFormattato(gfx, rowLay["eti_s011"].ToString(), prv, offsetX, offsetY, zIndexFilter);

            // [s014] ORIGINE
            DisegnaElementoTesto(gfx, rowLay["eti_s014"].ToString(), Convert.ToString(rowDati["eti_bor"]), offsetX, offsetY, zIndexFilter);

            // [s015] CALIBRO
            DisegnaElementoTesto(gfx, rowLay["eti_s015"].ToString(), Convert.ToString(rowDati["eti_bcl"]), offsetX, offsetY, zIndexFilter);

            // [s016] TASTO BILANCIA
            DisegnaElementoTesto(gfx, rowLay["eti_s016"].ToString(), Convert.ToString(rowDati["eti_tas"]), offsetX, offsetY, zIndexFilter);

            // [s018] UNITA DI MISURA
            DisegnaElementoTesto(gfx, rowLay["eti_s018"].ToString(), Convert.ToString(rowDati["eti_umi"]), offsetX, offsetY, zIndexFilter);

            // [s019] GRAMMATURA / PESO
            string s019val = "";
            if (!DBNull.Value.Equals(rowDati["eti_tgv"]) && Convert.ToDecimal(rowDati["eti_tgv"]) > 0)
                s019val = Convert.ToDecimal(rowDati["eti_tgv"]).ToString("0.###");
            DisegnaElementoTesto(gfx, rowLay["eti_s019"].ToString(), s019val, offsetX, offsetY, zIndexFilter);

            // [s020] PEZZI
            string s020val = "";
            if (!DBNull.Value.Equals(rowDati["eti_pxc"]) && Convert.ToDecimal(rowDati["eti_pxc"]) > 0)
                s020val = Convert.ToDecimal(rowDati["eti_pxc"]).ToString("0.#");
            DisegnaElementoTesto(gfx, rowLay["eti_s020"].ToString(), s020val, offsetX, offsetY, zIndexFilter);

            // [s021] MERCEOLOGIA
            DisegnaElementoTesto(gfx, rowLay["eti_s021"].ToString(), Convert.ToString(rowDati["eti_bct"]), offsetX, offsetY, zIndexFilter);

            // [s022] REPARTO
            string s022val = Convert.ToString(rowDati["eti_red"]);
            if (string.IsNullOrEmpty(s022val)) s022val = Convert.ToString(rowDati["eti_rep"]);
            DisegnaElementoTesto(gfx, rowLay["eti_s022"].ToString(), s022val, offsetX, offsetY, zIndexFilter);

            // [s023] DATA STAMPA
            if (rowLay.Table.Columns.Contains("eti_s023"))
                DisegnaElementoTesto(gfx, rowLay["eti_s023"].ToString(), DateTime.Now.ToString("dd/MM/yy"), offsetX, offsetY, zIndexFilter);

            // [s024] PERIODO PROMO
            if (rowLay.Table.Columns.Contains("eti_s024"))
            {
                string s024val = "";
                if (rowDati.Table.Columns.Contains("eti_odi") && !DBNull.Value.Equals(rowDati["eti_odi"]))
                {
                    DateTime d1 = Convert.ToDateTime(rowDati["eti_odi"]);
                    s024val = "Dal " + d1.ToString("dd/MM");
                    if (rowDati.Table.Columns.Contains("eti_odf") && !DBNull.Value.Equals(rowDati["eti_odf"]))
                    {
                        DateTime d2 = Convert.ToDateTime(rowDati["eti_odf"]);
                        s024val += " al " + d2.ToString("dd/MM");
                    }
                }
                DisegnaElementoTesto(gfx, rowLay["eti_s024"].ToString(), s024val, offsetX, offsetY, zIndexFilter);
            }

            // [s025] PREZZO UNITARIO Lt/Kg
            if (rowLay.Table.Columns.Contains("eti_s025"))
            {
                string s025val = "";
                if (!DBNull.Value.Equals(rowDati["eti_pne"]))
                {
                    decimal pne = Convert.ToDecimal(rowDati["eti_pne"]);
                    if (pne > 0 && pne <= 500m)
                    {
                        string umi = Convert.ToString(rowDati["eti_umi"]).ToLower();
                        if (umi == "lt") s025val = "€ " + pne.ToString("0.00") + "/lt";
                        else s025val = "€ " + pne.ToString("0.00") + "/kg";
                    }
                }
                DisegnaElementoTesto(gfx, rowLay["eti_s025"].ToString(), s025val, offsetX, offsetY, zIndexFilter);
            }

            // [eti_t001 .. eti_t008] TESTI LIBERI PERSONALIZZATI
            for (int i = 1; i <= 8; i++)
            {
                string colName = "eti_t00" + i;
                if (!rowLay.Table.Columns.Contains(colName)) continue;
                string rawCfg = rowLay[colName].ToString();
                if (string.IsNullOrEmpty(rawCfg) || !rawCfg.StartsWith("S")) continue;

                string[] parts = rawCfg.Split(';');
                string testoLibero = (parts.Length > 4) ? parts[4] : "";
                if (string.IsNullOrEmpty(testoLibero)) continue;

                DisegnaElementoTesto(gfx, rawCfg, testoLibero, offsetX, offsetY, zIndexFilter);
            }
        }
        private void DisegnaDescrizioneFit(XGraphics gfx, string cfg, string testo, double offsetX, double offsetY, int zIndexFilter = -1)
        {
            if (string.IsNullOrEmpty(cfg) || !cfg.StartsWith("S") || string.IsNullOrEmpty(testo)) return;

            // Filtro ZIndex
            if (zIndexFilter >= 0 && GetCfgZIndex(cfg) != zIndexFilter) return;

            string[] a = cfg.Split(';');
            if (a.Length < 3) return;

            string[] pos = a[1].Split(',');
            if (pos.Length < 4) return;

            double y = Convert.ToDouble(pos[0]) * 72.0 / 25.4 + offsetY;
            double x = Convert.ToDouble(pos[1]) * 72.0 / 25.4 + offsetX;
            double w = Convert.ToDouble(pos[2]) * 72.0 / 25.4;
            double h = Convert.ToDouble(pos[3]) * 72.0 / 25.4;

            XStringFormat format = new XStringFormat();
            if (pos.Length > 4 && pos[4] == "2") format.Alignment = XStringAlignment.Center;
            else if (pos.Length > 4 && pos[4] == "1") format.Alignment = XStringAlignment.Far;
            else format.Alignment = XStringAlignment.Near;
            format.LineAlignment = XLineAlignment.Center;

            string[] fnt = a[2].Split('|');
            string fName = fnt.Length > 0 && fnt[0] != "" ? fnt[0] : "Arial";
            double fSizeBase = fnt.Length > 1 && fnt[1] != "" ? Convert.ToDouble(fnt[1]) : 10;
            XFontStyle fStyle = (fnt.Length > 2 && fnt[2].ToUpper() == "B") ? XFontStyle.Bold : XFontStyle.Regular;

            XRect rect = new XRect(x, y, w, h);

            int fontColorArgb = Color.Black.ToArgb();
            if (a.Length > 3)
            {
                string[] shp = a[3].Split('|');
                if (shp.Length > 1) int.TryParse(shp[1], out fontColorArgb);
            }
            XSolidBrush fBrush = new XSolidBrush(XColor.FromArgb(fontColorArgb));

            // Tenta con font pieno a rimpicciolire
            for (double fTry = fSizeBase; fTry >= 4; fTry -= 0.5)
            {
                XFont font = new XFont(fName, fTry, fStyle);
                XSize size = gfx.MeasureString(testo, font);

                // Caso 1: entra tutto su una riga
                if (size.Width <= w)
                {
                    gfx.DrawString(testo, font, fBrush, rect, format);
                    return;
                }

                // Caso 2: proviamo a spezzare su 2 righe
                // Solo se l'altezza lo permette (almeno 2 righe teoriche)
                if (size.Height * 2 <= h + 4)
                {
                    string[] words = testo.Split(new char[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                    string line1 = "", line2 = "";
                    bool fits = false;
                    double minDiff = double.MaxValue;

                    for (int i = 1; i < words.Length; i++)
                    {
                        string r1 = string.Join(" ", words.Take(i).ToArray());
                        string r2 = string.Join(" ", words.Skip(i).ToArray());
                        XSize s1 = gfx.MeasureString(r1, font);
                        XSize s2 = gfx.MeasureString(r2, font);

                        if (s1.Width <= w && s2.Width <= w)
                        {
                            fits = true;
                            double diff = Math.Abs(s1.Width - s2.Width);
                            if (diff < minDiff)
                            {
                                minDiff = diff;
                                line1 = r1;
                                line2 = r2;
                            }
                        }
                    }

                    if (fits)
                    {
                        double lineH = h / 2.0;
                        gfx.DrawString(line1, font, fBrush, new XRect(x, y, w, lineH), format);
                        gfx.DrawString(line2, font, fBrush, new XRect(x, y + lineH, w, lineH), format);
                        return;
                    }
                }
            }

            // Fallback: se niente funziona stampa con font piccolissimo al centro
            XFont fallback = new XFont(fName, 4, fStyle);
            gfx.DrawString(testo, fallback, fBrush, rect, format);
        }


        private void DisegnaPrezzoFormattato(XGraphics gfx, string cfg, decimal prezzo, double offsetX, double offsetY, int zIndexFilter = -1)
        {
            if (string.IsNullOrEmpty(cfg) || !cfg.StartsWith("S")) return;

            // Filtro ZIndex
            if (zIndexFilter >= 0 && GetCfgZIndex(cfg) != zIndexFilter) return;

            string[] a = cfg.Split(';');
            if (a.Length < 3) return;

            string[] pos = a[1].Split(',');
            if (pos.Length < 4) return;

            double y = Convert.ToDouble(pos[0]) * 72.0 / 25.4 + offsetY;
            double x = Convert.ToDouble(pos[1]) * 72.0 / 25.4 + offsetX;
            double w = Convert.ToDouble(pos[2]) * 72.0 / 25.4;
            double h = Convert.ToDouble(pos[3]) * 72.0 / 25.4;

            string[] fnt = a[2].Split('|');
            string fName = fnt.Length > 0 && fnt[0] != "" ? fnt[0] : "Arial";
            double fSizeBase = fnt.Length > 1 && fnt[1] != "" ? Convert.ToDouble(fnt[1]) : 10;
            XFontStyle fStyle = (fnt.Length > 2 && fnt[2].ToUpper() == "B") ? XFontStyle.Bold : XFontStyle.Regular;

            int fontColorArgb = Color.Black.ToArgb();
            if (a.Length > 3)
            {
                string[] shp = a[3].Split('|');
                if (shp.Length > 1) int.TryParse(shp[1], out fontColorArgb);
            }
            XSolidBrush fBrush = new XSolidBrush(XColor.FromArgb(fontColorArgb));

            // Dimensioni: euro a 30% del font base, decimali a 50%
            double fSizeEuro = fSizeBase * 0.30;
            double fSizeDec = fSizeBase * 0.50;

            XFont fontEuro = new XFont(fName, fSizeEuro, fStyle);
            XFont fontMain = new XFont(fName, fSizeBase, fStyle);
            XFont fontDec = new XFont(fName, fSizeDec, fStyle);

            // Split prezzo: intero + decimali
            string intPart = ((int)prezzo).ToString();
            string decPart = (prezzo - (int)prezzo).ToString("0.00").Substring(1); // ",99"

            // Misura le parti per allineamento
            XSize sEuro = gfx.MeasureString("\u20ac", fontEuro);
            XSize sInt = gfx.MeasureString(intPart, fontMain);
            XSize sDec = gfx.MeasureString(decPart, fontDec);

            // Spazi e padding
            double spPadding = fontMain.GetHeight() * 0.35; // Spazio largo "prima" e "dopo"
            double spEuroInt = fontMain.GetHeight() * 0.15; // Spazio tra Euro e Intero
            double spIntDec = fontMain.GetHeight() * 0.05;  // Spazio tra Intero e Decimale
            double totalW = (spPadding * 2) + sEuro.Width + spEuroInt + sInt.Width + spIntDec + sDec.Width;

            // Centratura orizzontale
            double xStart = x + (w - totalW) / 2.0;
            if (xStart < x) xStart = x;
            xStart += spPadding;

            XStringFormat fmt = new XStringFormat();
            fmt.Alignment = XStringAlignment.Near;
            fmt.LineAlignment = XLineAlignment.Near;

            // Allineamento verticale
            double yBase = y + h - fontMain.GetHeight();
            if (yBase < y) yBase = y;

            // Euro allineato in alto (linea della part intera)
            double yEuro = yBase; // Nessun disallineamento negativo, parte con in linea
            double yInt = yBase;
            double yDec = yBase + (fontMain.GetHeight() * 0.15);

            gfx.DrawString("\u20ac", fontEuro, fBrush, new XRect(xStart, yEuro, sEuro.Width + 5, h), fmt);
            xStart += sEuro.Width + spEuroInt;

            gfx.DrawString(intPart, fontMain, fBrush, new XRect(xStart, yInt, sInt.Width + 5, h), fmt);
            xStart += sInt.Width + spIntDec;

            gfx.DrawString(decPart, fontDec, fBrush, new XRect(xStart, yDec, sDec.Width + 5, h), fmt);
        }
    }
}
