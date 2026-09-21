using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace APOffice
{
    public static class clsUiIcons
    {
        public static Bitmap GetIcon(string iconName, int size = 16, Color? primaryColor = null, Color? accentColor = null)
        {
            Bitmap bmp = new Bitmap(size, size);
            using (Graphics g = Graphics.FromImage(bmp))
            {
                g.SmoothingMode = SmoothingMode.AntiAlias;
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;

                Color pri = primaryColor ?? Color.FromArgb(51, 65, 85);       // Slate 700
                Color acc = accentColor ?? Color.FromArgb(37, 99, 235);       // Blue 600
                Color danger = Color.FromArgb(220, 38, 38);                    // Red 600
                Color success = Color.FromArgb(22, 163, 74);                   // Green 600
                Color warning = Color.FromArgb(217, 119, 6);                   // Amber 600

                float s = size / 16.0f;

                switch (iconName.ToLowerInvariant())
                {
                    case "exit":
                    case "close":
                        // Porta con freccia uscita
                        using (Pen p = new Pen(danger, 1.5f * s))
                        {
                            g.DrawRectangle(p, 2 * s, 2 * s, 6 * s, 12 * s);
                            g.DrawLine(p, 6 * s, 8 * s, 14 * s, 8 * s);
                            g.DrawLine(p, 11 * s, 5 * s, 14 * s, 8 * s);
                            g.DrawLine(p, 11 * s, 11 * s, 14 * s, 8 * s);
                        }
                        break;

                    case "plus":
                    case "add":
                    case "nuovo":
                        // Segno più / aggiungi
                        using (Pen p = new Pen(primaryColor ?? success, 1.8f * s))
                        {
                            g.DrawLine(p, 8 * s, 3 * s, 8 * s, 13 * s);
                            g.DrawLine(p, 3 * s, 8 * s, 13 * s, 8 * s);
                        }
                        break;

                    case "copy":
                    case "duplicate":
                    case "duplica":
                        // Due fogli sovrapposti
                        using (Pen p = new Pen(pri, 1.2f * s))
                        using (SolidBrush bBg1 = new SolidBrush(Color.FromArgb(241, 245, 249)))
                        using (SolidBrush bBg2 = new SolidBrush(Color.White))
                        {
                            g.FillRectangle(bBg1, 5 * s, 2 * s, 8 * s, 10 * s);
                            g.DrawRectangle(p, 5 * s, 2 * s, 8 * s, 10 * s);
                            g.FillRectangle(bBg2, 2 * s, 5 * s, 8 * s, 10 * s);
                            g.DrawRectangle(p, 2 * s, 5 * s, 8 * s, 10 * s);
                        }
                        break;

                    case "save":
                    case "salva":
                    case "disk":
                        // Floppy disk salvataggio
                        using (Pen p = new Pen(pri, 1.2f * s))
                        using (SolidBrush b = new SolidBrush(Color.FromArgb(241, 245, 249)))
                        using (SolidBrush bShutter = new SolidBrush(Color.FromArgb(148, 163, 184)))
                        {
                            g.FillRectangle(b, 2.5f * s, 2.5f * s, 11 * s, 11 * s);
                            g.DrawRectangle(p, 2.5f * s, 2.5f * s, 11 * s, 11 * s);
                            g.FillRectangle(bShutter, 4.5f * s, 3 * s, 7 * s, 4 * s);
                            g.DrawRectangle(p, 4.5f * s, 3 * s, 7 * s, 4 * s);
                            g.FillRectangle(bShutter, 4 * s, 9 * s, 8 * s, 4.5f * s);
                        }
                        break;

                    case "receipt":
                    case "scontrino":
                        // Rotolo scontrino con linee e dentellatura
                        using (Pen p = new Pen(pri, 1.2f * s))
                        using (SolidBrush b = new SolidBrush(Color.FromArgb(241, 245, 249)))
                        {
                            PointF[] pts = new PointF[] {
                                new PointF(3 * s, 2 * s),
                                new PointF(13 * s, 2 * s),
                                new PointF(13 * s, 14 * s),
                                new PointF(11 * s, 12.5f * s),
                                new PointF(9 * s, 14 * s),
                                new PointF(7 * s, 12.5f * s),
                                new PointF(5 * s, 14 * s),
                                new PointF(3 * s, 12.5f * s)
                            };
                            g.FillPolygon(b, pts);
                            g.DrawPolygon(p, pts);
                            // Righe di testo
                            g.DrawLine(p, 5 * s, 5 * s, 11 * s, 5 * s);
                            g.DrawLine(p, 5 * s, 7.5f * s, 11 * s, 7.5f * s);
                            g.DrawLine(p, 5 * s, 10 * s, 9 * s, 10 * s);
                        }
                        break;

                    case "delete_rows":
                    case "eraser":
                        // Foglio con cancellazione/meno
                        using (Pen p = new Pen(pri, 1.2f * s))
                        using (Pen pd = new Pen(danger, 1.6f * s))
                        {
                            g.DrawRectangle(p, 2.5f * s, 2.5f * s, 11 * s, 11 * s);
                            g.DrawLine(p, 4.5f * s, 5.5f * s, 11.5f * s, 5.5f * s);
                            g.DrawLine(pd, 4.5f * s, 8.5f * s, 11.5f * s, 8.5f * s);
                            g.DrawLine(p, 4.5f * s, 11.5f * s, 9.5f * s, 11.5f * s);
                        }
                        break;

                    case "trash":
                    case "delete":
                        // Cestino dei rifiuti
                        using (Pen p = new Pen(danger, 1.4f * s))
                        using (SolidBrush b = new SolidBrush(Color.FromArgb(254, 226, 226)))
                        {
                            g.FillRectangle(b, 4 * s, 5 * s, 8 * s, 9 * s);
                            g.DrawRectangle(p, 4 * s, 5 * s, 8 * s, 9 * s);
                            g.DrawLine(p, 2 * s, 5 * s, 14 * s, 5 * s);
                            g.DrawLine(p, 6 * s, 3 * s, 10 * s, 3 * s);
                            g.DrawLine(p, 6.5f * s, 7 * s, 6.5f * s, 12 * s);
                            g.DrawLine(p, 9.5f * s, 7 * s, 9.5f * s, 12 * s);
                        }
                        break;

                    case "terminal":
                    case "barcode":
                        // Terminalino / Scanner barcode
                        using (Pen p = new Pen(pri, 1.2f * s))
                        using (SolidBrush b = new SolidBrush(acc))
                        {
                            DrawRoundedRectangleHelper(g, p, 3 * s, 2 * s, 10 * s, 12 * s, 2 * s);
                            g.FillRectangle(b, 5 * s, 3.5f * s, 6 * s, 4 * s);
                            // Tasti
                            g.DrawLine(p, 5 * s, 9.5f * s, 7 * s, 9.5f * s);
                            g.DrawLine(p, 9 * s, 9.5f * s, 11 * s, 9.5f * s);
                            g.DrawLine(p, 5 * s, 12 * s, 7 * s, 12 * s);
                            g.DrawLine(p, 9 * s, 12 * s, 11 * s, 12 * s);
                        }
                        break;

                    case "import":
                    case "download":
                        // Cartella / Vassoio con freccia import
                        using (Pen p = new Pen(acc, 1.5f * s))
                        using (Pen pb = new Pen(pri, 1.2f * s))
                        {
                            g.DrawLine(pb, 2 * s, 10 * s, 2 * s, 13.5f * s);
                            g.DrawLine(pb, 2 * s, 13.5f * s, 14 * s, 13.5f * s);
                            g.DrawLine(pb, 14 * s, 13.5f * s, 14 * s, 10 * s);

                            g.DrawLine(p, 8 * s, 2.5f * s, 8 * s, 10.5f * s);
                            g.DrawLine(p, 5 * s, 7.5f * s, 8 * s, 10.5f * s);
                            g.DrawLine(p, 11 * s, 7.5f * s, 8 * s, 10.5f * s);
                        }
                        break;

                    case "document":
                    case "invoice":
                        // Documento / Fattura con piega
                        using (Pen p = new Pen(pri, 1.2f * s))
                        using (SolidBrush b = new SolidBrush(Color.FromArgb(248, 250, 252)))
                        using (SolidBrush bh = new SolidBrush(acc))
                        {
                            PointF[] docPts = new PointF[] {
                                new PointF(3 * s, 2 * s),
                                new PointF(10 * s, 2 * s),
                                new PointF(13 * s, 5 * s),
                                new PointF(13 * s, 14 * s),
                                new PointF(3 * s, 14 * s)
                            };
                            g.FillPolygon(b, docPts);
                            g.DrawPolygon(p, docPts);
                            g.FillRectangle(bh, 5 * s, 4.5f * s, 4 * s, 2 * s);
                            g.DrawLine(p, 5 * s, 8.5f * s, 11 * s, 8.5f * s);
                            g.DrawLine(p, 5 * s, 11 * s, 10 * s, 11 * s);
                        }
                        break;

                    case "xml":
                    case "e_invoice":
                    case "sdi":
                        // Fattura elettronica XML con tag < / > e icona documento
                        using (Pen p = new Pen(Color.FromArgb(30, 64, 175), 1.2f * s))
                        using (SolidBrush b = new SolidBrush(Color.FromArgb(238, 242, 255)))
                        {
                            FillRoundedRectangleHelper(g, b, 2 * s, 2 * s, 12 * s, 12 * s, 2 * s);
                            DrawRoundedRectangleHelper(g, p, 2 * s, 2 * s, 12 * s, 12 * s, 2 * s);

                            // Tag < >
                            using (Pen pTag = new Pen(Color.FromArgb(79, 70, 229), 1.6f * s))
                            {
                                g.DrawLine(pTag, 6 * s, 5 * s, 4 * s, 8 * s);
                                g.DrawLine(pTag, 4 * s, 8 * s, 6 * s, 11 * s);

                                g.DrawLine(pTag, 10 * s, 5 * s, 12 * s, 8 * s);
                                g.DrawLine(pTag, 12 * s, 8 * s, 10 * s, 11 * s);

                                g.DrawLine(pTag, 9 * s, 5 * s, 7 * s, 11 * s);
                            }
                        }
                        break;

                    case "scale":
                    case "bilancia":
                        // Bilancia a piatti
                        using (Pen p = new Pen(pri, 1.2f * s))
                        {
                            g.DrawLine(p, 8 * s, 3 * s, 8 * s, 14 * s);
                            g.DrawLine(p, 4 * s, 14 * s, 12 * s, 14 * s);
                            g.DrawLine(p, 3 * s, 5 * s, 13 * s, 5 * s);
                            // Piatti
                            g.DrawLine(p, 3 * s, 5 * s, 2 * s, 9 * s);
                            g.DrawLine(p, 3 * s, 5 * s, 4 * s, 9 * s);
                            g.DrawLine(p, 1.5f * s, 9 * s, 4.5f * s, 9 * s);

                            g.DrawLine(p, 13 * s, 5 * s, 12 * s, 9 * s);
                            g.DrawLine(p, 13 * s, 5 * s, 14 * s, 9 * s);
                            g.DrawLine(p, 11.5f * s, 9 * s, 14.5f * s, 9 * s);
                        }
                        break;

                    case "chart":
                    case "stats":
                        // Istogramma con freccia di crescita
                        using (SolidBrush b1 = new SolidBrush(Color.FromArgb(148, 163, 184)))
                        using (SolidBrush b2 = new SolidBrush(Color.FromArgb(96, 165, 250)))
                        using (SolidBrush b3 = new SolidBrush(acc))
                        using (Pen p = new Pen(success, 1.4f * s))
                        {
                            g.FillRectangle(b1, 3 * s, 9 * s, 2.5f * s, 5 * s);
                            g.FillRectangle(b2, 6.5f * s, 6 * s, 2.5f * s, 8 * s);
                            g.FillRectangle(b3, 10 * s, 3 * s, 2.5f * s, 11 * s);
                            // Trend line
                            g.DrawLine(p, 2 * s, 10 * s, 7 * s, 5 * s);
                            g.DrawLine(p, 7 * s, 5 * s, 13 * s, 2 * s);
                        }
                        break;

                    case "check":
                    case "verify":
                        // Spunta di verifica
                        using (Pen p = new Pen(success, 1.8f * s))
                        {
                            g.DrawLine(p, 3 * s, 8 * s, 6.5f * s, 11.5f * s);
                            g.DrawLine(p, 6.5f * s, 11.5f * s, 13 * s, 4 * s);
                        }
                        break;

                    case "print":
                        // Stampante moderna
                        using (Pen p = new Pen(pri, 1.2f * s))
                        using (SolidBrush b = new SolidBrush(Color.FromArgb(241, 245, 249)))
                        using (SolidBrush bPaper = new SolidBrush(Color.White))
                        {
                            g.FillRectangle(bPaper, 4.5f * s, 2 * s, 7 * s, 4 * s);
                            g.DrawRectangle(p, 4.5f * s, 2 * s, 7 * s, 4 * s);
                            g.FillRectangle(b, 2 * s, 5 * s, 12 * s, 6 * s);
                            g.DrawRectangle(p, 2 * s, 5 * s, 12 * s, 6 * s);
                            g.FillRectangle(bPaper, 4 * s, 8.5f * s, 8 * s, 5 * s);
                            g.DrawRectangle(p, 4 * s, 8.5f * s, 8 * s, 5 * s);
                            g.DrawLine(p, 5.5f * s, 10.5f * s, 10.5f * s, 10.5f * s);
                        }
                        break;

                    case "search":
                    case "find":
                        // Lente d'ingrandimento
                        using (Pen p = new Pen(pri, 1.5f * s))
                        {
                            g.DrawEllipse(p, 3 * s, 3 * s, 7 * s, 7 * s);
                            g.DrawLine(p, 8.5f * s, 8.5f * s, 13.5f * s, 13.5f * s);
                        }
                        break;

                    case "user":
                    case "client":
                        // Utente / Cliente
                        using (SolidBrush b = new SolidBrush(acc))
                        {
                            g.FillEllipse(b, 5 * s, 2.5f * s, 6 * s, 6 * s);
                            g.FillPie(b, 2 * s, 7.5f * s, 12 * s, 11 * s, 180, 180);
                        }
                        break;

                    case "box":
                    case "lotti":
                        // Scatola / Lotto
                        using (Pen p = new Pen(warning, 1.3f * s))
                        using (SolidBrush b = new SolidBrush(Color.FromArgb(254, 243, 199)))
                        {
                            PointF[] bPts = new PointF[] {
                                new PointF(8 * s, 2 * s),
                                new PointF(14 * s, 5 * s),
                                new PointF(8 * s, 8 * s),
                                new PointF(2 * s, 5 * s)
                            };
                            g.FillPolygon(b, bPts);
                            g.DrawPolygon(p, bPts);
                            g.DrawLine(p, 2 * s, 5 * s, 2 * s, 11.5f * s);
                            g.DrawLine(p, 14 * s, 5 * s, 14 * s, 11.5f * s);
                            g.DrawLine(p, 8 * s, 8 * s, 8 * s, 14.5f * s);
                            g.DrawLine(p, 2 * s, 11.5f * s, 8 * s, 14.5f * s);
                            g.DrawLine(p, 14 * s, 11.5f * s, 8 * s, 14.5f * s);
                        }
                        break;

                    case "discount":
                    case "percent":
                        // Badge percentuale sconto
                        using (Pen p = new Pen(Color.FromArgb(225, 29, 72), 1.3f * s))
                        using (SolidBrush b = new SolidBrush(Color.FromArgb(255, 228, 230)))
                        {
                            g.FillEllipse(b, 2 * s, 2 * s, 12 * s, 12 * s);
                            g.DrawEllipse(p, 2 * s, 2 * s, 12 * s, 12 * s);
                            g.DrawLine(p, 10.5f * s, 5.5f * s, 5.5f * s, 10.5f * s);
                            using (SolidBrush bDot = new SolidBrush(Color.FromArgb(225, 29, 72)))
                            {
                                g.FillEllipse(bDot, 5 * s, 5 * s, 2 * s, 2 * s);
                                g.FillEllipse(bDot, 9 * s, 9 * s, 2 * s, 2 * s);
                            }
                        }
                        break;

                    case "refresh":
                    case "sync":
                        // Frecce circolari aggiorna
                        using (Pen p = new Pen(acc, 1.4f * s))
                        {
                            g.DrawArc(p, 3 * s, 3 * s, 10 * s, 10 * s, 45, 270);
                            g.DrawLine(p, 10 * s, 2 * s, 13 * s, 5 * s);
                            g.DrawLine(p, 13 * s, 5 * s, 10 * s, 8 * s);
                        }
                        break;

                    case "send":
                        // Freccia invio
                        using (SolidBrush b = new SolidBrush(success))
                        {
                            PointF[] sPts = new PointF[] {
                                new PointF(2 * s, 2 * s),
                                new PointF(14 * s, 8 * s),
                                new PointF(2 * s, 14 * s),
                                new PointF(5 * s, 8 * s)
                            };
                            g.FillPolygon(b, sPts);
                        }
                        break;

                    case "articles":
                    case "products":
                        // Box merce 3D con codice a barre e nastro
                        using (Pen pBox = new Pen(pri, 1.4f * s))
                        using (SolidBrush bTop = new SolidBrush(Color.FromArgb(224, 231, 255)))
                        using (SolidBrush bSide = new SolidBrush(Color.FromArgb(199, 210, 254)))
                        using (Pen pLine = new Pen(acc, 1.2f * s))
                        {
                            PointF[] topPts = new PointF[] {
                                new PointF(8 * s, 1.5f * s),
                                new PointF(14.5f * s, 4.5f * s),
                                new PointF(8 * s, 7.5f * s),
                                new PointF(1.5f * s, 4.5f * s)
                            };
                            g.FillPolygon(bTop, topPts);
                            g.DrawPolygon(pBox, topPts);

                            PointF[] leftPts = new PointF[] {
                                new PointF(1.5f * s, 4.5f * s),
                                new PointF(8 * s, 7.5f * s),
                                new PointF(8 * s, 14.5f * s),
                                new PointF(1.5f * s, 11.5f * s)
                            };
                            g.FillPolygon(bSide, leftPts);
                            g.DrawPolygon(pBox, leftPts);

                            PointF[] rightPts = new PointF[] {
                                new PointF(8 * s, 7.5f * s),
                                new PointF(14.5f * s, 4.5f * s),
                                new PointF(14.5f * s, 11.5f * s),
                                new PointF(8 * s, 14.5f * s)
                            };
                            g.FillPolygon(bTop, rightPts);
                            g.DrawPolygon(pBox, rightPts);

                            // Barcode decorativo sul lato sinistro
                            g.DrawLine(pLine, 3.5f * s, 7.5f * s, 3.5f * s, 11.5f * s);
                            g.DrawLine(pLine, 5.0f * s, 8.2f * s, 5.0f * s, 12.2f * s);
                            g.DrawLine(pLine, 6.5f * s, 9.0f * s, 6.5f * s, 13.0f * s);
                        }
                        break;

                    case "suppliers":
                    case "truck":
                        // Camion di fornitura / logistica con ruote e cabina
                        using (Pen pTrk = new Pen(pri, 1.3f * s))
                        using (SolidBrush bBody = new SolidBrush(Color.FromArgb(209, 250, 229)))
                        using (SolidBrush bCab = new SolidBrush(Color.FromArgb(167, 243, 208)))
                        using (SolidBrush bWheel = new SolidBrush(Color.FromArgb(30, 41, 59)))
                        {
                            // Cassone
                            g.FillRectangle(bBody, 1.5f * s, 3.5f * s, 8.5f * s, 7.5f * s);
                            g.DrawRectangle(pTrk, 1.5f * s, 3.5f * s, 8.5f * s, 7.5f * s);

                            // Cabina
                            PointF[] cabPts = new PointF[] {
                                new PointF(10.0f * s, 5.5f * s),
                                new PointF(13.0f * s, 5.5f * s),
                                new PointF(14.5f * s, 8.0f * s),
                                new PointF(14.5f * s, 11.0f * s),
                                new PointF(10.0f * s, 11.0f * s)
                            };
                            g.FillPolygon(bCab, cabPts);
                            g.DrawPolygon(pTrk, cabPts);

                            // Finestrino
                            using (SolidBrush bWin = new SolidBrush(Color.FromArgb(240, 253, 250)))
                            {
                                g.FillRectangle(bWin, 10.5f * s, 6.5f * s, 2.5f * s, 2.5f * s);
                                g.DrawRectangle(pTrk, 10.5f * s, 6.5f * s, 2.5f * s, 2.5f * s);
                            }

                            // Ruote
                            g.FillEllipse(bWheel, 3.0f * s, 10.0f * s, 3.5f * s, 3.5f * s);
                            g.DrawEllipse(pTrk, 3.0f * s, 10.0f * s, 3.5f * s, 3.5f * s);

                            g.FillEllipse(bWheel, 10.5f * s, 10.0f * s, 3.5f * s, 3.5f * s);
                            g.DrawEllipse(pTrk, 10.5f * s, 10.0f * s, 3.5f * s, 3.5f * s);
                        }
                        break;

                    case "documents_stack":
                    case "docs_multi":
                        // Pila di fatture / documenti con pieghe e linee
                        using (Pen pDoc = new Pen(pri, 1.2f * s))
                        using (SolidBrush bBg1 = new SolidBrush(Color.FromArgb(254, 243, 199)))
                        using (SolidBrush bBg2 = new SolidBrush(Color.FromArgb(255, 251, 235)))
                        {
                            // Foglio posteriore
                            g.FillRectangle(bBg1, 4.5f * s, 1.5f * s, 9.5f * s, 11.5f * s);
                            g.DrawRectangle(pDoc, 4.5f * s, 1.5f * s, 9.5f * s, 11.5f * s);

                            // Foglio principale anteriore
                            PointF[] dPts = new PointF[] {
                                new PointF(2.0f * s, 3.5f * s),
                                new PointF(9.0f * s, 3.5f * s),
                                new PointF(12.0f * s, 6.5f * s),
                                new PointF(12.0f * s, 14.5f * s),
                                new PointF(2.0f * s, 14.5f * s)
                            };
                            g.FillPolygon(bBg2, dPts);
                            g.DrawPolygon(pDoc, dPts);

                            // Righe documento
                            using (Pen pLines = new Pen(pri, 1.1f * s))
                            {
                                g.DrawLine(pLines, 4.0f * s, 6.5f * s, 8.0f * s, 6.5f * s);
                                g.DrawLine(pLines, 4.0f * s, 9.0f * s, 10.0f * s, 9.0f * s);
                                g.DrawLine(pLines, 4.0f * s, 11.5f * s, 9.0f * s, 11.5f * s);
                            }
                        }
                        break;

                    case "clients_group":
                    case "users_group":
                        // Gruppo clienti (3 persone / profili)
                        using (SolidBrush bHead = new SolidBrush(pri))
                        using (SolidBrush bBody = new SolidBrush(Color.FromArgb(254, 205, 211)))
                        using (Pen pUser = new Pen(pri, 1.2f * s))
                        {
                            // Utente centrale
                            g.FillEllipse(bHead, 5.5f * s, 2.0f * s, 5.0f * s, 5.0f * s);
                            g.DrawEllipse(pUser, 5.5f * s, 2.0f * s, 5.0f * s, 5.0f * s);
                            g.FillPie(bBody, 2.5f * s, 7.5f * s, 11.0f * s, 9.0f * s, 180, 180);
                            g.DrawArc(pUser, 2.5f * s, 7.5f * s, 11.0f * s, 9.0f * s, 180, 180);

                            // Utente sinistro
                            g.FillEllipse(bHead, 1.5f * s, 4.5f * s, 3.5f * s, 3.5f * s);
                            g.DrawArc(pUser, 0.5f * s, 9.0f * s, 6.0f * s, 6.0f * s, 180, 180);

                            // Utente destro
                            g.FillEllipse(bHead, 11.0f * s, 4.5f * s, 3.5f * s, 3.5f * s);
                            g.DrawArc(pUser, 9.5f * s, 9.0f * s, 6.0f * s, 6.0f * s, 180, 180);
                        }
                        break;

                    case "fidelity_card":
                    case "tessere":
                        // Carta fedeltà con chip dorato e banda magnetica
                        using (Pen pCard = new Pen(pri, 1.3f * s))
                        using (SolidBrush bCard = new SolidBrush(Color.FromArgb(237, 233, 254)))
                        using (SolidBrush bChip = new SolidBrush(Color.FromArgb(234, 179, 8)))
                        using (SolidBrush bStripe = new SolidBrush(Color.FromArgb(109, 40, 217)))
                        {
                            DrawRoundedRectangleHelper(g, pCard, 1.5f * s, 3.5f * s, 13.0f * s, 9.5f * s, 2.0f * s);
                            FillRoundedRectangleHelper(g, bCard, 1.5f * s, 3.5f * s, 13.0f * s, 9.5f * s, 2.0f * s);

                            // Chip
                            g.FillRectangle(bChip, 3.5f * s, 6.0f * s, 3.0f * s, 2.2f * s);
                            g.DrawRectangle(pCard, 3.5f * s, 6.0f * s, 3.0f * s, 2.2f * s);

                            // Banda o linee
                            g.FillRectangle(bStripe, 1.5f * s, 9.5f * s, 13.0f * s, 1.8f * s);

                            // Stella fedeltà
                            using (SolidBrush bStar = new SolidBrush(Color.FromArgb(168, 85, 247)))
                            {
                                g.FillEllipse(bStar, 10.5f * s, 5.5f * s, 2.5f * s, 2.5f * s);
                            }
                        }
                        break;

                    case "cloud_sync":
                    case "divulgazioni":
                        // Nuvola con sincronizzazione / download listini
                        using (Pen pCloud = new Pen(pri, 1.3f * s))
                        using (SolidBrush bCloud = new SolidBrush(Color.FromArgb(204, 251, 241)))
                        using (Pen pSync = new Pen(acc, 1.5f * s))
                        {
                            // Profilo nuvola
                            g.FillEllipse(bCloud, 2.0f * s, 4.5f * s, 5.5f * s, 5.5f * s);
                            g.FillEllipse(bCloud, 5.5f * s, 2.0f * s, 6.0f * s, 6.0f * s);
                            g.FillEllipse(bCloud, 9.5f * s, 4.0f * s, 4.5f * s, 4.5f * s);
                            g.FillRectangle(bCloud, 3.5f * s, 6.0f * s, 9.0f * s, 4.5f * s);

                            g.DrawArc(pCloud, 2.0f * s, 4.5f * s, 5.5f * s, 5.5f * s, 140, 180);
                            g.DrawArc(pCloud, 5.5f * s, 2.0f * s, 6.0f * s, 6.0f * s, 190, 160);
                            g.DrawArc(pCloud, 9.5f * s, 4.0f * s, 4.5f * s, 4.5f * s, 260, 170);
                            g.DrawLine(pCloud, 3.5f * s, 10.5f * s, 12.5f * s, 10.5f * s);

                            // Frecce circolari di sync
                            g.DrawArc(pSync, 5.0f * s, 9.5f * s, 6.0f * s, 5.0f * s, 30, 260);
                            g.DrawLine(pSync, 10.5f * s, 10.0f * s, 11.5f * s, 12.5f * s);
                            g.DrawLine(pSync, 9.0f * s, 12.0f * s, 11.5f * s, 12.5f * s);
                        }
                        break;

                    case "orders_cart":
                    case "cart":
                        // Carrello ordini e acquisti
                        using (Pen pCart = new Pen(pri, 1.4f * s))
                        using (SolidBrush bWheel = new SolidBrush(Color.FromArgb(30, 41, 59)))
                        using (SolidBrush bBasket = new SolidBrush(Color.FromArgb(224, 242, 254)))
                        {
                            // Manico e base
                            g.DrawLine(pCart, 1.5f * s, 2.5f * s, 4.0f * s, 2.5f * s);
                            g.DrawLine(pCart, 4.0f * s, 2.5f * s, 5.5f * s, 9.5f * s);
                            g.DrawLine(pCart, 5.5f * s, 9.5f * s, 13.5f * s, 9.5f * s);
                            g.DrawLine(pCart, 13.5f * s, 9.5f * s, 14.5f * s, 4.5f * s);
                            g.DrawLine(pCart, 14.5f * s, 4.5f * s, 4.5f * s, 4.5f * s);

                            // Griglia carrello
                            using (Pen pGrid = new Pen(acc, 1.0f * s))
                            {
                                g.DrawLine(pGrid, 7.5f * s, 4.5f * s, 7.0f * s, 9.5f * s);
                                g.DrawLine(pGrid, 10.5f * s, 4.5f * s, 10.0f * s, 9.5f * s);
                                g.DrawLine(pGrid, 5.0f * s, 7.0f * s, 14.0f * s, 7.0f * s);
                            }

                            // Ruote
                            g.FillEllipse(bWheel, 5.5f * s, 11.0f * s, 2.8f * s, 2.8f * s);
                            g.FillEllipse(bWheel, 11.5f * s, 11.0f * s, 2.8f * s, 2.8f * s);
                        }
                        break;

                    case "basket":
                    case "shopping_basket":
                    case "paniere":
                        // Cestino della spesa / Paniere intrecciato con manico
                        using (Pen pBsk = new Pen(pri, 1.3f * s))
                        using (SolidBrush bBsk = new SolidBrush(Color.FromArgb(254, 243, 199)))
                        {
                            // Manico superiore
                            g.DrawArc(pBsk, 4.0f * s, 1.5f * s, 8.0f * s, 6.0f * s, 180, 180);

                            // Cestino
                            PointF[] bPts = new PointF[] {
                                new PointF(2.0f * s, 6.0f * s),
                                new PointF(14.0f * s, 6.0f * s),
                                new PointF(12.0f * s, 13.5f * s),
                                new PointF(4.0f * s, 13.5f * s)
                            };
                            g.FillPolygon(bBsk, bPts);
                            g.DrawPolygon(pBsk, bPts);

                            // Intreccio
                            using (Pen pWeave = new Pen(acc, 1.0f * s))
                            {
                                g.DrawLine(pWeave, 5.5f * s, 6.0f * s, 6.0f * s, 13.5f * s);
                                g.DrawLine(pWeave, 8.0f * s, 6.0f * s, 8.0f * s, 13.5f * s);
                                g.DrawLine(pWeave, 10.5f * s, 6.0f * s, 10.0f * s, 13.5f * s);
                                g.DrawLine(pWeave, 2.8f * s, 9.5f * s, 13.2f * s, 9.5f * s);
                            }
                        }
                        break;

                    case "offers_promo":
                    case "promotions":
                        // Cartellino offerta con percentuale % e nastro promo
                        using (Pen pOff = new Pen(pri, 1.3f * s))
                        using (SolidBrush bTag = new SolidBrush(Color.FromArgb(255, 237, 213)))
                        {
                            PointF[] tagPts = new PointF[] {
                                new PointF(2.5f * s, 7.0f * s),
                                new PointF(7.0f * s, 2.5f * s),
                                new PointF(13.5f * s, 9.0f * s),
                                new PointF(9.0f * s, 13.5f * s)
                            };
                            g.FillPolygon(bTag, tagPts);
                            g.DrawPolygon(pOff, tagPts);

                            // Foro cartellino
                            using (SolidBrush bHole = new SolidBrush(Color.White))
                            {
                                g.FillEllipse(bHole, 5.0f * s, 4.5f * s, 2.2f * s, 2.2f * s);
                                g.DrawEllipse(pOff, 5.0f * s, 4.5f * s, 2.2f * s, 2.2f * s);
                            }

                            // % Simbolo
                            using (Pen pPct = new Pen(pri, 1.3f * s))
                            using (SolidBrush bPct = new SolidBrush(pri))
                            {
                                g.DrawLine(pPct, 11.0f * s, 7.0f * s, 8.0f * s, 11.0f * s);
                                g.FillEllipse(bPct, 7.5f * s, 7.0f * s, 1.8f * s, 1.8f * s);
                                g.FillEllipse(bPct, 10.0f * s, 10.0f * s, 1.8f * s, 1.8f * s);
                            }
                        }
                        break;

                    case "inventory_warehouse":
                    case "inventory":
                        // Cartella inventario con spunte e scatola
                        using (Pen pInv = new Pen(pri, 1.3f * s))
                        using (SolidBrush bBoard = new SolidBrush(Color.FromArgb(224, 242, 254)))
                        using (Pen pChk = new Pen(Color.FromArgb(14, 116, 144), 1.4f * s))
                        {
                            DrawRoundedRectangleHelper(g, pInv, 2.5f * s, 2.5f * s, 11.0f * s, 12.0f * s, 2.0f * s);
                            FillRoundedRectangleHelper(g, bBoard, 2.5f * s, 2.5f * s, 11.0f * s, 12.0f * s, 2.0f * s);

                            // Clip in alto
                            using (SolidBrush bClip = new SolidBrush(pri))
                            {
                                g.FillRectangle(bClip, 5.5f * s, 1.5f * s, 5.0f * s, 2.5f * s);
                            }

                            // Righe con spunte
                            g.DrawLine(pChk, 4.5f * s, 6.0f * s, 5.8f * s, 7.5f * s);
                            g.DrawLine(pChk, 5.8f * s, 7.5f * s, 7.5f * s, 5.0f * s);
                            g.DrawLine(pInv, 8.5f * s, 6.0f * s, 11.5f * s, 6.0f * s);

                            g.DrawLine(pChk, 4.5f * s, 9.5f * s, 5.8f * s, 11.0f * s);
                            g.DrawLine(pChk, 5.8f * s, 11.0f * s, 7.5f * s, 8.5f * s);
                            g.DrawLine(pInv, 8.5f * s, 9.5f * s, 11.5f * s, 9.5f * s);
                        }
                        break;

                    case "pos_balance":
                    case "variazioni":
                        // Cassa POS e Bilancia elettronica
                        using (Pen pPos = new Pen(pri, 1.3f * s))
                        using (SolidBrush bPos = new SolidBrush(Color.FromArgb(238, 242, 255)))
                        using (SolidBrush bSc = new SolidBrush(Color.FromArgb(99, 102, 241)))
                        {
                            // Monitor POS
                            g.FillRectangle(bPos, 2.0f * s, 2.0f * s, 8.0f * s, 6.5f * s);
                            g.DrawRectangle(pPos, 2.0f * s, 2.0f * s, 8.0f * s, 6.5f * s);
                            g.FillRectangle(bSc, 3.2f * s, 3.2f * s, 5.6f * s, 4.0f * s);

                            // Stand e tastiera
                            g.DrawLine(pPos, 6.0f * s, 8.5f * s, 6.0f * s, 10.5f * s);
                            g.DrawLine(pPos, 2.0f * s, 10.5f * s, 10.0f * s, 10.5f * s);

                            // Piatto bilancia
                            g.DrawLine(pPos, 10.5f * s, 7.5f * s, 14.5f * s, 7.5f * s);
                            g.DrawLine(pPos, 12.5f * s, 7.5f * s, 12.5f * s, 13.5f * s);
                            g.DrawLine(pPos, 1.5f * s, 13.5f * s, 14.5f * s, 13.5f * s);
                        }
                        break;

                    case "day_closing":
                    case "chiusura":
                        // Orologio fine turno con lucchetto
                        using (Pen pClk = new Pen(pri, 1.3f * s))
                        using (SolidBrush bClk = new SolidBrush(Color.FromArgb(241, 245, 249)))
                        using (SolidBrush bLock = new SolidBrush(Color.FromArgb(100, 116, 139)))
                        {
                            g.FillEllipse(bClk, 1.5f * s, 1.5f * s, 11.5f * s, 11.5f * s);
                            g.DrawEllipse(pClk, 1.5f * s, 1.5f * s, 11.5f * s, 11.5f * s);

                            // Lancette
                            g.DrawLine(pClk, 7.2f * s, 7.2f * s, 7.2f * s, 3.5f * s);
                            g.DrawLine(pClk, 7.2f * s, 7.2f * s, 10.0f * s, 7.2f * s);

                            // Lucchetto in basso a destra
                            g.FillRectangle(bLock, 9.0f * s, 8.5f * s, 5.5f * s, 5.0f * s);
                            g.DrawRectangle(pClk, 9.0f * s, 8.5f * s, 5.5f * s, 5.0f * s);
                            g.DrawArc(pClk, 10.0f * s, 6.0f * s, 3.5f * s, 4.0f * s, 180, 180);
                        }
                        break;

                    case "excel_import":
                    case "excel":
                    case "xls":
                        // Icona foglio Excel verde con "X" e griglia dati
                        using (Pen pXls = new Pen(Color.FromArgb(21, 128, 61), 1.2f * s))
                        using (SolidBrush bXls = new SolidBrush(Color.FromArgb(220, 252, 231)))
                        using (SolidBrush bBadge = new SolidBrush(Color.FromArgb(22, 163, 74)))
                        {
                            // Cartella foglio
                            g.FillRectangle(bXls, 3.5f * s, 2.0f * s, 10.5f * s, 12.0f * s);
                            g.DrawRectangle(pXls, 3.5f * s, 2.0f * s, 10.5f * s, 12.0f * s);

                            // Griglia interna
                            using (Pen pGrid = new Pen(Color.FromArgb(134, 239, 172), 1.0f * s))
                            {
                                g.DrawLine(pGrid, 7.5f * s, 2.0f * s, 7.5f * s, 14.0f * s);
                                g.DrawLine(pGrid, 10.5f * s, 2.0f * s, 10.5f * s, 14.0f * s);
                                g.DrawLine(pGrid, 3.5f * s, 5.5f * s, 14.0f * s, 5.5f * s);
                                g.DrawLine(pGrid, 3.5f * s, 9.0f * s, 14.0f * s, 9.0f * s);
                            }

                            // Badge "X" verde scuro
                            g.FillRectangle(bBadge, 1.5f * s, 4.5f * s, 6.0f * s, 7.0f * s);
                            using (Pen pX = new Pen(Color.White, 1.6f * s))
                            {
                                g.DrawLine(pX, 3.0f * s, 6.0f * s, 6.0f * s, 10.0f * s);
                                g.DrawLine(pX, 6.0f * s, 6.0f * s, 3.0f * s, 10.0f * s);
                            }
                        }
                        break;

                    case "tasks_prints":
                    case "attivita":
                    case "stampe":
                        // Stampante con report / lista attività
                        using (Pen pPrn = new Pen(pri, 1.3f * s))
                        using (SolidBrush bPrn = new SolidBrush(Color.FromArgb(254, 249, 195)))
                        using (SolidBrush bPaper = new SolidBrush(Color.White))
                        {
                            // Carta superiore
                            g.FillRectangle(bPaper, 4.5f * s, 1.5f * s, 7.0f * s, 4.0f * s);
                            g.DrawRectangle(pPrn, 4.5f * s, 1.5f * s, 7.0f * s, 4.0f * s);

                            // Corpo stampante
                            g.FillRectangle(bPrn, 2.0f * s, 5.0f * s, 12.0f * s, 5.5f * s);
                            g.DrawRectangle(pPrn, 2.0f * s, 5.0f * s, 12.0f * s, 5.5f * s);

                            // Foglio in uscita
                            g.FillRectangle(bPaper, 4.0f * s, 8.5f * s, 8.0f * s, 5.5f * s);
                            g.DrawRectangle(pPrn, 4.0f * s, 8.5f * s, 8.0f * s, 5.5f * s);

                            using (Pen pL = new Pen(pri, 1.0f * s))
                            {
                                g.DrawLine(pL, 5.5f * s, 10.5f * s, 10.5f * s, 10.5f * s);
                                g.DrawLine(pL, 5.5f * s, 12.5f * s, 9.0f * s, 12.5f * s);
                            }
                        }
                        break;

                    case "statistics":
                    case "stats_graph":
                        // Grafico statistico 3 barre con freccia di salita
                        using (SolidBrush b1 = new SolidBrush(Color.FromArgb(147, 197, 253)))
                        using (SolidBrush b2 = new SolidBrush(Color.FromArgb(96, 165, 250)))
                        using (SolidBrush b3 = new SolidBrush(Color.FromArgb(37, 99, 235)))
                        using (Pen pTrend = new Pen(Color.FromArgb(22, 163, 74), 1.8f * s))
                        {
                            g.FillRectangle(b1, 2.5f * s, 8.5f * s, 3.0f * s, 5.5f * s);
                            g.FillRectangle(b2, 6.5f * s, 5.5f * s, 3.0f * s, 8.5f * s);
                            g.FillRectangle(b3, 10.5f * s, 2.5f * s, 3.0f * s, 11.5f * s);

                            // Linea trend e freccia
                            g.DrawLine(pTrend, 1.5f * s, 10.5f * s, 7.5f * s, 5.0f * s);
                            g.DrawLine(pTrend, 7.5f * s, 5.0f * s, 13.5f * s, 1.5f * s);
                            g.DrawLine(pTrend, 10.5f * s, 1.5f * s, 13.5f * s, 1.5f * s);
                            g.DrawLine(pTrend, 13.5f * s, 1.5f * s, 13.5f * s, 4.5f * s);
                        }
                        break;

                    case "tools_gear":
                    case "utilita":
                        // Chiave inglese e ingranaggio impostazioni
                        using (Pen pTool = new Pen(pri, 1.4f * s))
                        using (SolidBrush bGear = new SolidBrush(Color.FromArgb(226, 232, 240)))
                        {
                            // Ingranaggio
                            g.FillEllipse(bGear, 3.5f * s, 3.5f * s, 9.0f * s, 9.0f * s);
                            g.DrawEllipse(pTool, 3.5f * s, 3.5f * s, 9.0f * s, 9.0f * s);
                            g.DrawEllipse(pTool, 6.0f * s, 6.0f * s, 4.0f * s, 4.0f * s);

                            // Denti ingranaggio
                            g.DrawLine(pTool, 8.0f * s, 1.5f * s, 8.0f * s, 3.5f * s);
                            g.DrawLine(pTool, 8.0f * s, 12.5f * s, 8.0f * s, 14.5f * s);
                            g.DrawLine(pTool, 1.5f * s, 8.0f * s, 3.5f * s, 8.0f * s);
                            g.DrawLine(pTool, 12.5f * s, 8.0f * s, 14.5f * s, 8.0f * s);

                            // Chiave inglese diagonale
                            using (Pen pWrench = new Pen(Color.FromArgb(71, 85, 105), 1.8f * s))
                            {
                                g.DrawLine(pWrench, 2.5f * s, 13.5f * s, 11.5f * s, 4.5f * s);
                            }
                        }
                        break;

                    case "ai_agent":
                    case "ai":
                    case "robot":
                        // Assistente AI con scintille / robot
                        using (Pen pAi = new Pen(Color.FromArgb(124, 58, 237), 1.3f * s))
                        using (SolidBrush bAi = new SolidBrush(Color.FromArgb(245, 243, 255)))
                        {
                            DrawRoundedRectangleHelper(g, pAi, 3.0f * s, 4.0f * s, 10.0f * s, 8.5f * s, 2.5f * s);
                            FillRoundedRectangleHelper(g, bAi, 3.0f * s, 4.0f * s, 10.0f * s, 8.5f * s, 2.5f * s);

                            // Occhi visore
                            using (SolidBrush bEye = new SolidBrush(Color.FromArgb(124, 58, 237)))
                            {
                                g.FillEllipse(bEye, 5.0f * s, 6.5f * s, 2.2f * s, 2.2f * s);
                                g.FillEllipse(bEye, 8.8f * s, 6.5f * s, 2.2f * s, 2.2f * s);
                            }

                            // Antenna
                            g.DrawLine(pAi, 8.0f * s, 4.0f * s, 8.0f * s, 1.5f * s);
                            using (SolidBrush bDot = new SolidBrush(Color.FromArgb(234, 179, 8)))
                            {
                                g.FillEllipse(bDot, 7.0f * s, 1.0f * s, 2.0f * s, 2.0f * s);
                            }
                        }
                        break;

                    case "labels":
                    case "etichette":
                        // Rotolo etichette adesive barcode
                        using (Pen pEti = new Pen(pri, 1.3f * s))
                        using (SolidBrush bEti = new SolidBrush(Color.FromArgb(240, 253, 244)))
                        {
                            g.FillRectangle(bEti, 2.0f * s, 3.0f * s, 12.0f * s, 10.0f * s);
                            g.DrawRectangle(pEti, 2.0f * s, 3.0f * s, 12.0f * s, 10.0f * s);

                            // Barcode
                            using (Pen pB = new Pen(pri, 1.2f * s))
                            {
                                g.DrawLine(pB, 4.0f * s, 5.5f * s, 4.0f * s, 10.5f * s);
                                g.DrawLine(pB, 6.0f * s, 5.5f * s, 6.0f * s, 10.5f * s);
                                g.DrawLine(pB, 7.5f * s, 5.5f * s, 7.5f * s, 10.5f * s);
                                g.DrawLine(pB, 9.5f * s, 5.5f * s, 9.5f * s, 10.5f * s);
                                g.DrawLine(pB, 11.5f * s, 5.5f * s, 11.5f * s, 10.5f * s);
                            }
                        }
                        break;

                    case "department":
                    case "reparto":
                        // Scaffalatura / Reparto negozio
                        using (Pen pDep = new Pen(pri, 1.3f * s))
                        using (SolidBrush bRoof = new SolidBrush(Color.FromArgb(224, 242, 254)))
                        {
                            // Tettuccio reparto
                            PointF[] rPts = new PointF[] {
                                new PointF(1.5f * s, 5.0f * s),
                                new PointF(8.0f * s, 2.0f * s),
                                new PointF(14.5f * s, 5.0f * s)
                            };
                            g.FillPolygon(bRoof, rPts);
                            g.DrawPolygon(pDep, rPts);

                            // Colonne e scaffale
                            g.DrawLine(pDep, 3.0f * s, 5.0f * s, 3.0f * s, 13.5f * s);
                            g.DrawLine(pDep, 13.0f * s, 5.0f * s, 13.0f * s, 13.5f * s);
                            g.DrawLine(pDep, 2.0f * s, 9.0f * s, 14.0f * s, 9.0f * s);
                            g.DrawLine(pDep, 1.5f * s, 13.5f * s, 14.5f * s, 13.5f * s);

                            // Merci su scaffale
                            using (SolidBrush bGoods = new SolidBrush(acc))
                            {
                                g.FillRectangle(bGoods, 4.5f * s, 6.5f * s, 3.0f * s, 2.5f * s);
                                g.FillRectangle(bGoods, 8.5f * s, 6.5f * s, 3.0f * s, 2.5f * s);
                                g.FillRectangle(bGoods, 4.5f * s, 10.5f * s, 7.0f * s, 3.0f * s);
                            }
                        }
                        break;

                    case "categories":
                    case "merceologia":
                        // Cartellini categorie e merceologia
                        using (Pen pCat = new Pen(pri, 1.3f * s))
                        using (SolidBrush bBgCat = new SolidBrush(Color.FromArgb(204, 251, 241)))
                        {
                            PointF[] cPts = new PointF[] {
                                new PointF(2.0f * s, 4.0f * s),
                                new PointF(8.5f * s, 4.0f * s),
                                new PointF(14.0f * s, 9.5f * s),
                                new PointF(8.5f * s, 14.5f * s),
                                new PointF(2.0f * s, 9.0f * s)
                            };
                            g.FillPolygon(bBgCat, cPts);
                            g.DrawPolygon(pCat, cPts);

                            using (SolidBrush bHole = new SolidBrush(Color.White))
                            {
                                g.FillEllipse(bHole, 4.5f * s, 6.0f * s, 2.5f * s, 2.5f * s);
                                g.DrawEllipse(pCat, 4.5f * s, 6.0f * s, 2.5f * s, 2.5f * s);
                            }
                            g.DrawLine(pCat, 8.0f * s, 8.5f * s, 11.5f * s, 12.0f * s);
                        }
                        break;

                    case "calendar_week":
                    case "settimanali":
                        // Calendario estrazioni settimanali
                        using (Pen pCal = new Pen(pri, 1.2f * s))
                        using (SolidBrush bCal = new SolidBrush(Color.FromArgb(224, 242, 254)))
                        using (SolidBrush bHdr = new SolidBrush(Color.FromArgb(14, 116, 144)))
                        {
                            DrawRoundedRectangleHelper(g, pCal, 2.0f * s, 2.5f * s, 12.0f * s, 11.5f * s, 2.0f * s);
                            FillRoundedRectangleHelper(g, bCal, 2.0f * s, 2.5f * s, 12.0f * s, 11.5f * s, 2.0f * s);

                            g.FillRectangle(bHdr, 2.0f * s, 2.5f * s, 12.0f * s, 3.5f * s);

                            // Griglia giorni
                            using (SolidBrush bDot = new SolidBrush(pri))
                            {
                                g.FillEllipse(bDot, 4.0f * s, 8.0f * s, 1.5f * s, 1.5f * s);
                                g.FillEllipse(bDot, 7.5f * s, 8.0f * s, 1.5f * s, 1.5f * s);
                                g.FillEllipse(bDot, 11.0f * s, 8.0f * s, 1.5f * s, 1.5f * s);
                                g.FillEllipse(bDot, 4.0f * s, 11.0f * s, 1.5f * s, 1.5f * s);
                                g.FillEllipse(bDot, 7.5f * s, 11.0f * s, 1.5f * s, 1.5f * s);
                                g.FillEllipse(bDot, 11.0f * s, 11.0f * s, 1.5f * s, 1.5f * s);
                            }
                        }
                        break;

                    case "gift_award":
                    case "premi":
                        // Pacco regalo con fiocco per premi fidelity
                        using (Pen pGift = new Pen(pri, 1.3f * s))
                        using (SolidBrush bBox = new SolidBrush(Color.FromArgb(237, 233, 254)))
                        using (SolidBrush bRibbon = new SolidBrush(Color.FromArgb(147, 51, 234)))
                        {
                            // Coperchio
                            g.FillRectangle(bBox, 2.0f * s, 4.5f * s, 12.0f * s, 3.0f * s);
                            g.DrawRectangle(pGift, 2.0f * s, 4.5f * s, 12.0f * s, 3.0f * s);

                            // Base scatola
                            g.FillRectangle(bBox, 3.0f * s, 7.5f * s, 10.0f * s, 6.5f * s);
                            g.DrawRectangle(pGift, 3.0f * s, 7.5f * s, 10.0f * s, 6.5f * s);

                            // Nastro
                            g.FillRectangle(bRibbon, 7.0f * s, 4.5f * s, 2.0f * s, 9.5f * s);

                            // Fiocco in cima
                            g.DrawArc(pGift, 4.0f * s, 1.5f * s, 4.0f * s, 3.5f * s, 180, 240);
                            g.DrawArc(pGift, 8.0f * s, 1.5f * s, 4.0f * s, 3.5f * s, 120, 240);
                        }
                        break;

                    case "celiachia":
                    case "gluten_free":
                        // Simbolo spiga sbarrata / alimentazione speciale
                        using (Pen pHeart = new Pen(Color.FromArgb(225, 29, 72), 1.4f * s))
                        using (SolidBrush bBgH = new SolidBrush(Color.FromArgb(255, 228, 230)))
                        {
                            g.FillEllipse(bBgH, 2.0f * s, 2.0f * s, 12.0f * s, 12.0f * s);
                            g.DrawEllipse(pHeart, 2.0f * s, 2.0f * s, 12.0f * s, 12.0f * s);

                            // Croce medica / spiga
                            using (SolidBrush bCross = new SolidBrush(Color.FromArgb(225, 29, 72)))
                            {
                                g.FillRectangle(bCross, 6.5f * s, 4.5f * s, 3.0f * s, 7.0f * s);
                                g.FillRectangle(bCross, 4.5f * s, 6.5f * s, 7.0f * s, 3.0f * s);
                            }
                        }
                        break;

                    case "history_clock":
                    case "storico":
                        // Orologio con freccia all'indietro (storico)
                        using (Pen pHist = new Pen(pri, 1.4f * s))
                        {
                            g.DrawArc(pHist, 2.5f * s, 2.5f * s, 11.0f * s, 11.0f * s, 45, 280);
                            // Freccia rewind
                            PointF[] arrPts = new PointF[] {
                                new PointF(10.0f * s, 1.5f * s),
                                new PointF(13.5f * s, 3.5f * s),
                                new PointF(10.0f * s, 5.5f * s)
                            };
                            using (SolidBrush bArr = new SolidBrush(pri))
                            {
                                g.FillPolygon(bArr, arrPts);
                            }
                            // Lancette
                            g.DrawLine(pHist, 8.0f * s, 8.0f * s, 8.0f * s, 5.0f * s);
                            g.DrawLine(pHist, 8.0f * s, 8.0f * s, 10.5f * s, 8.0f * s);
                        }
                        break;

                    case "extract":
                    case "filtra":
                        // Lente filtro / estrazione dati
                        using (Pen pExt = new Pen(Color.White, 1.5f * s))
                        {
                            g.DrawEllipse(pExt, 3.0f * s, 3.0f * s, 6.5f * s, 6.5f * s);
                            g.DrawLine(pExt, 8.0f * s, 8.0f * s, 13.0f * s, 13.0f * s);
                            g.DrawLine(pExt, 4.5f * s, 6.2f * s, 8.0f * s, 6.2f * s);
                        }
                        break;

                    default:
                        using (Pen p = new Pen(pri, 1.2f * s))
                        {
                            g.DrawRectangle(p, 3 * s, 3 * s, 10 * s, 10 * s);
                        }
                        break;
                }
            }
            return bmp;
        }

        private class MainMenuTileState
        {
            public int BadgeNum;
            public string Title;
            public string IconName;
            public int IconSize;
            public Color GradTop;
            public Color GradBottom;
            public Color BorderCol;
            public Color TextCol;
            public Color IconCol;
            public bool IsHovered;
            public bool IsPressed;
        }

        private static GraphicsPath CreateRoundedTilePath(Rectangle rect, int radius)
        {
            GraphicsPath path = new GraphicsPath();
            int d = radius * 2;
            if (d > rect.Width) d = rect.Width;
            if (d > rect.Height) d = rect.Height;

            path.AddArc(rect.X, rect.Y, d, d, 180, 90);
            path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90);
            path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90);
            path.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }

        public static void StyleMainMenuButton(Button btn, int badgeNum, string title, string iconName, int iconSize,
            Color gradTop, Color gradBottom, Color borderCol, Color textCol, Color iconCol)
        {
            if (btn == null) return;
            btn.Text = "";
            btn.Image = null;
            btn.BackgroundImage = null;
            btn.FlatStyle = FlatStyle.Flat;
            btn.FlatAppearance.BorderSize = 0;
            btn.BackColor = gradTop;
            btn.Cursor = Cursors.Hand;

            // Rimuovi eventuali label residue create in precedenza
            List<Control> toRemove = new List<Control>();
            foreach (Control sub in btn.Controls)
            {
                if (sub.Tag != null && sub.Tag.ToString() == "BtnNum")
                    toRemove.Add(sub);
            }
            foreach (Control sub in toRemove)
                btn.Controls.Remove(sub);

            MainMenuTileState state = btn.Tag as MainMenuTileState;
            if (state == null)
            {
                state = new MainMenuTileState();
                btn.Tag = state;

                btn.MouseEnter += (s, e) => { state.IsHovered = true; btn.Invalidate(); };
                btn.MouseLeave += (s, e) => { state.IsHovered = false; state.IsPressed = false; btn.Invalidate(); };
                btn.MouseDown += (s, e) => { if (e.Button == MouseButtons.Left) { state.IsPressed = true; btn.Invalidate(); } };
                btn.MouseUp += (s, e) => { state.IsPressed = false; btn.Invalidate(); };

                btn.Paint += (s, e) =>
                {
                    MainMenuTileState st = btn.Tag as MainMenuTileState;
                    if (st == null) return;

                    Graphics g = e.Graphics;
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

                    Rectangle drawRect = new Rectangle(0, 0, btn.Width - 1, btn.Height - 1);

                    Color cTop = st.GradTop;
                    Color cBottom = st.GradBottom;

                    if (st.IsHovered && !st.IsPressed)
                    {
                        // Effetto WOW: gradiente luminoso e brillante al passaggio mouse
                        cTop = Color.FromArgb(
                            Math.Min(255, cTop.R + 18),
                            Math.Min(255, cTop.G + 18),
                            Math.Min(255, cTop.B + 18));
                        cBottom = Color.FromArgb(
                            Math.Min(255, cBottom.R + 25),
                            Math.Min(255, cBottom.G + 25),
                            Math.Min(255, cBottom.B + 25));
                    }
                    else if (st.IsPressed)
                    {
                        // Effetto feedback click
                        cTop = Color.FromArgb(
                            Math.Max(0, cTop.R - 25),
                            Math.Max(0, cTop.G - 25),
                            Math.Max(0, cTop.B - 25));
                        cBottom = Color.FromArgb(
                            Math.Max(0, cBottom.R - 20),
                            Math.Max(0, cBottom.G - 20),
                            Math.Max(0, cBottom.B - 20));
                    }

                    // Riempimento con gradiente sfumato multi-stop lucido
                    using (GraphicsPath path = CreateRoundedTilePath(drawRect, 6))
                    {
                        using (LinearGradientBrush lgb = new LinearGradientBrush(drawRect, cTop, cBottom, LinearGradientMode.Vertical))
                        {
                            ColorBlend cb = new ColorBlend(3);
                            cb.Colors = new Color[] {
                                cTop,
                                Color.FromArgb((cTop.R * 2 + cBottom.R) / 3, (cTop.G * 2 + cBottom.G) / 3, (cTop.B * 2 + cBottom.B) / 3),
                                cBottom
                            };
                            cb.Positions = new float[] { 0.0f, 0.40f, 1.0f };
                            lgb.InterpolationColors = cb;

                            g.FillPath(lgb, path);
                        }

                        // Riflesso superiore lucido (Shine)
                        using (Pen pShine = new Pen(Color.FromArgb(170, 255, 255, 255), 1.0f))
                        {
                            g.DrawLine(pShine, drawRect.Left + 5, drawRect.Top + 1, drawRect.Right - 5, drawRect.Top + 1);
                        }

                        // Bordo coordinato
                        Color bColor = st.IsHovered ? st.IconCol : st.BorderCol;
                        using (Pen pBorder = new Pen(bColor, st.IsHovered ? 1.5f : 1.0f))
                        {
                            g.DrawPath(pBorder, path);
                        }
                    }

                    // 1. Disegno Badge Numerico in alto a sinistra (sempre visibile e proporzionato)
                    if (st.BadgeNum > 0)
                    {
                        int badgeW = (st.BadgeNum >= 10) ? 25 : 20;
                        int badgeH = 18;
                        Rectangle badgeRect = new Rectangle(5, 5, badgeW, badgeH);

                        using (GraphicsPath bPath = CreateRoundedTilePath(badgeRect, 3))
                        using (SolidBrush bBg = new SolidBrush(Color.FromArgb(30, 41, 59))) // Slate 800
                        using (SolidBrush bText = new SolidBrush(Color.White))
                        using (Font bFont = new Font("Segoe UI", 8.0f, FontStyle.Bold))
                        using (StringFormat sfBadge = new StringFormat() { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center })
                        {
                            g.FillPath(bBg, bPath);
                            g.DrawString(st.BadgeNum.ToString(), bFont, bText, badgeRect, sfBadge);
                        }
                    }

                    // 2. Disegno Icona vettoriale
                    int isz = st.IconSize;
                    int iconX = (btn.Width >= 300) ? 42 : ((btn.Height >= 70) ? 34 : 34);
                    int iconY = (btn.Height - isz) / 2;

                    using (Bitmap ico = GetIcon(st.IconName, isz, st.IconCol, st.IconCol))
                    {
                        if (ico != null)
                        {
                            g.DrawImage(ico, iconX, iconY, isz, isz);
                        }
                    }

                    // 3. Disegno Testo del Tasto
                    int textX = iconX + isz + 10;
                    int textW = btn.Width - textX - 6;
                    int textH = btn.Height - 4;
                    Rectangle textRect = new Rectangle(textX, 2, textW, textH);

                    float fontSize = (btn.Height >= 70 && btn.Width >= 300) ? 10.5f : ((btn.Height >= 70) ? 9.5f : 9.0f);
                    using (Font tFont = new Font("Segoe UI", fontSize, FontStyle.Bold))
                    using (SolidBrush tBrush = new SolidBrush(st.TextCol))
                    using (StringFormat sfText = new StringFormat() { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Center, FormatFlags = StringFormatFlags.NoClip })
                    {
                        g.DrawString(st.Title, tFont, tBrush, textRect, sfText);
                    }
                };
            }

            state.BadgeNum = badgeNum;
            state.Title = title;
            state.IconName = iconName;
            state.IconSize = iconSize;
            state.GradTop = gradTop;
            state.GradBottom = gradBottom;
            state.BorderCol = borderCol;
            state.TextCol = textCol;
            state.IconCol = iconCol;

            btn.Invalidate();
        }

        public static void StyleStatButton(Button btn, string title, string iconName, int iconSize,
            Color gradTop, Color gradBottom, Color borderCol, Color textCol, Color iconCol)
        {
            if (btn == null) return;
            btn.Text = "";
            btn.Image = null;
            btn.BackgroundImage = null;
            btn.FlatStyle = FlatStyle.Flat;
            btn.FlatAppearance.BorderSize = 0;
            btn.BackColor = gradTop;
            btn.Cursor = Cursors.Hand;

            MainMenuTileState state = btn.Tag as MainMenuTileState;
            if (state == null)
            {
                state = new MainMenuTileState();
                btn.Tag = state;

                btn.MouseEnter += (s, e) => { state.IsHovered = true; btn.Invalidate(); };
                btn.MouseLeave += (s, e) => { state.IsHovered = false; state.IsPressed = false; btn.Invalidate(); };
                btn.MouseDown += (s, e) => { if (e.Button == MouseButtons.Left) { state.IsPressed = true; btn.Invalidate(); } };
                btn.MouseUp += (s, e) => { state.IsPressed = false; btn.Invalidate(); };

                btn.Paint += (s, e) =>
                {
                    MainMenuTileState st = btn.Tag as MainMenuTileState;
                    if (st == null) return;

                    Graphics g = e.Graphics;
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

                    Rectangle drawRect = new Rectangle(0, 0, btn.Width - 1, btn.Height - 1);

                    Color cTop = st.GradTop;
                    Color cBottom = st.GradBottom;

                    if (st.IsHovered && !st.IsPressed)
                    {
                        cTop = Color.FromArgb(
                            Math.Min(255, cTop.R + 18),
                            Math.Min(255, cTop.G + 18),
                            Math.Min(255, cTop.B + 18));
                        cBottom = Color.FromArgb(
                            Math.Min(255, cBottom.R + 25),
                            Math.Min(255, cBottom.G + 25),
                            Math.Min(255, cBottom.B + 25));
                    }
                    else if (st.IsPressed)
                    {
                        cTop = Color.FromArgb(
                            Math.Max(0, cTop.R - 25),
                            Math.Max(0, cTop.G - 25),
                            Math.Max(0, cTop.B - 25));
                        cBottom = Color.FromArgb(
                            Math.Max(0, cBottom.R - 20),
                            Math.Max(0, cBottom.G - 20),
                            Math.Max(0, cBottom.B - 20));
                    }

                    // Riempimento con gradiente lucido
                    using (GraphicsPath path = CreateRoundedTilePath(drawRect, 6))
                    {
                        using (LinearGradientBrush lgb = new LinearGradientBrush(drawRect, cTop, cBottom, LinearGradientMode.Vertical))
                        {
                            ColorBlend cb = new ColorBlend(3);
                            cb.Colors = new Color[] {
                                cTop,
                                Color.FromArgb((cTop.R * 2 + cBottom.R) / 3, (cTop.G * 2 + cBottom.G) / 3, (cTop.B * 2 + cBottom.B) / 3),
                                cBottom
                            };
                            cb.Positions = new float[] { 0.0f, 0.40f, 1.0f };
                            lgb.InterpolationColors = cb;

                            g.FillPath(lgb, path);
                        }

                        // Riflesso superiore
                        using (Pen pShine = new Pen(Color.FromArgb(170, 255, 255, 255), 1.0f))
                        {
                            g.DrawLine(pShine, drawRect.Left + 5, drawRect.Top + 1, drawRect.Right - 5, drawRect.Top + 1);
                        }

                        // Bordo
                        Color bColor = st.IsHovered ? st.IconCol : st.BorderCol;
                        using (Pen pBorder = new Pen(bColor, st.IsHovered ? 1.5f : 1.0f))
                        {
                            g.DrawPath(pBorder, path);
                        }
                    }

                    // Disegno Icona vettoriale
                    int isz = st.IconSize;
                    int iconX = 12;
                    int iconY = (btn.Height - isz) / 2;

                    using (Bitmap ico = GetIcon(st.IconName, isz, st.IconCol, st.IconCol))
                    {
                        if (ico != null)
                        {
                            g.DrawImage(ico, iconX, iconY, isz, isz);
                        }
                    }

                    // Disegno Testo del Tasto
                    int textX = iconX + isz + 8;
                    int textW = btn.Width - textX - 6;
                    int textH = btn.Height - 4;
                    Rectangle textRect = new Rectangle(textX, 2, textW, textH);

                    using (Font tFont = new Font("Segoe UI", 9.0f, FontStyle.Bold))
                    using (SolidBrush tBrush = new SolidBrush(st.TextCol))
                    using (StringFormat sfText = new StringFormat() { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Center, FormatFlags = StringFormatFlags.NoClip })
                    {
                        g.DrawString(st.Title, tFont, tBrush, textRect, sfText);
                    }
                };
            }

            state.BadgeNum = 0;
            state.Title = title;
            state.IconName = iconName;
            state.IconSize = iconSize;
            state.GradTop = gradTop;
            state.GradBottom = gradBottom;
            state.BorderCol = borderCol;
            state.TextCol = textCol;
            state.IconCol = iconCol;

            btn.Invalidate();
        }

        private class CustomImageTileState
        {
            public string Title;
            public Image CustomIcon;
            public Color GradTop;
            public Color GradBottom;
            public Color BorderCol;
            public Color TextCol;
            public Color HoverBorderCol;
            public bool IsHovered;
            public bool IsPressed;
        }

        public static void StyleUtilityButtonWithImage(Button btn, Image icon, Color gradTop, Color gradBottom,
            Color borderCol, Color textCol, Color hoverBorderCol)
        {
            if (btn == null) return;
            string originalText = !string.IsNullOrEmpty(btn.Text) ? btn.Text : (btn.Tag is CustomImageTileState st0 ? st0.Title : "");
            Image originalImage = icon ?? btn.Image ?? (btn.Tag is CustomImageTileState st1 ? st1.CustomIcon : null);

            // Mantieni sempre il testo originale su btn.Text per consentire ai gestori di eventi di leggerlo
            btn.Text = originalText;
            btn.Image = null;
            btn.BackgroundImage = null;
            btn.FlatStyle = FlatStyle.Flat;
            btn.FlatAppearance.BorderSize = 0;
            btn.BackColor = gradTop;
            btn.Cursor = Cursors.Hand;

            CustomImageTileState state = btn.Tag as CustomImageTileState;
            if (state == null)
            {
                state = new CustomImageTileState();
                btn.Tag = state;

                btn.MouseEnter += (s, e) => { state.IsHovered = true; btn.Invalidate(); };
                btn.MouseLeave += (s, e) => { state.IsHovered = false; state.IsPressed = false; btn.Invalidate(); };
                btn.MouseDown += (s, e) => { if (e.Button == MouseButtons.Left) { state.IsPressed = true; btn.Invalidate(); } };
                btn.MouseUp += (s, e) => { state.IsPressed = false; btn.Invalidate(); };

                btn.Paint += (s, e) =>
                {
                    CustomImageTileState st = btn.Tag as CustomImageTileState;
                    if (st == null) return;

                    Graphics g = e.Graphics;
                    g.SmoothingMode = SmoothingMode.AntiAlias;
                    g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                    g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.ClearTypeGridFit;

                    Rectangle drawRect = new Rectangle(0, 0, btn.Width - 1, btn.Height - 1);

                    Color cTop = st.GradTop;
                    Color cBottom = st.GradBottom;

                    if (st.IsHovered && !st.IsPressed)
                    {
                        cTop = Color.FromArgb(
                            Math.Min(255, cTop.R + 18),
                            Math.Min(255, cTop.G + 18),
                            Math.Min(255, cTop.B + 18));
                        cBottom = Color.FromArgb(
                            Math.Min(255, cBottom.R + 25),
                            Math.Min(255, cBottom.G + 25),
                            Math.Min(255, cBottom.B + 25));
                    }
                    else if (st.IsPressed)
                    {
                        cTop = Color.FromArgb(
                            Math.Max(0, cTop.R - 25),
                            Math.Max(0, cTop.G - 25),
                            Math.Max(0, cTop.B - 25));
                        cBottom = Color.FromArgb(
                            Math.Max(0, cBottom.R - 20),
                            Math.Max(0, cBottom.G - 20),
                            Math.Max(0, cBottom.B - 20));
                    }

                    // Riempimento con gradiente lucido
                    using (GraphicsPath path = CreateRoundedTilePath(drawRect, 6))
                    {
                        using (LinearGradientBrush lgb = new LinearGradientBrush(drawRect, cTop, cBottom, LinearGradientMode.Vertical))
                        {
                            ColorBlend cb = new ColorBlend(3);
                            cb.Colors = new Color[] {
                                cTop,
                                Color.FromArgb((cTop.R * 2 + cBottom.R) / 3, (cTop.G * 2 + cBottom.G) / 3, (cTop.B * 2 + cBottom.B) / 3),
                                cBottom
                            };
                            cb.Positions = new float[] { 0.0f, 0.40f, 1.0f };
                            lgb.InterpolationColors = cb;

                            g.FillPath(lgb, path);
                        }

                        // Riflesso superiore
                        using (Pen pShine = new Pen(Color.FromArgb(170, 255, 255, 255), 1.0f))
                        {
                            g.DrawLine(pShine, drawRect.Left + 5, drawRect.Top + 1, drawRect.Right - 5, drawRect.Top + 1);
                        }

                        // Bordo
                        Color bColor = st.IsHovered ? st.HoverBorderCol : st.BorderCol;
                        using (Pen pBorder = new Pen(bColor, st.IsHovered ? 1.5f : 1.0f))
                        {
                            g.DrawPath(pBorder, path);
                        }
                    }

                    // Disegno Icona esistente
                    int iconX = 8;
                    if (st.CustomIcon != null)
                    {
                        int imgW = Math.Min(st.CustomIcon.Width, 24);
                        int imgH = Math.Min(st.CustomIcon.Height, 24);
                        int imgY = (btn.Height - imgH) / 2;
                        g.DrawImage(st.CustomIcon, iconX, imgY, imgW, imgH);
                        iconX += imgW + 6;
                    }
                    else
                    {
                        iconX += 10;
                    }

                    // Disegno Testo del Tasto
                    string dispText = !string.IsNullOrEmpty(st.Title) ? st.Title : btn.Text;
                    int textW = btn.Width - iconX - 4;
                    int textH = btn.Height - 4;
                    Rectangle textRect = new Rectangle(iconX, 2, textW, textH);

                    using (Font tFont = new Font("Segoe UI", 8.5f, FontStyle.Bold))
                    using (SolidBrush tBrush = new SolidBrush(st.TextCol))
                    using (StringFormat sfText = new StringFormat() { Alignment = StringAlignment.Near, LineAlignment = StringAlignment.Center, FormatFlags = StringFormatFlags.NoClip })
                    {
                        g.DrawString(dispText, tFont, tBrush, textRect, sfText);
                    }
                };
            }

            state.Title = originalText;
            state.CustomIcon = originalImage;
            state.GradTop = gradTop;
            state.GradBottom = gradBottom;
            state.BorderCol = borderCol;
            state.TextCol = textCol;
            state.HoverBorderCol = hoverBorderCol;

            btn.Invalidate();
        }

        public static void StyleDataGridView(DataGridView dgv)
        {
            if (dgv == null) return;

            try
            {
                // Abilita DoubleBuffered per scorrimento fluido
                typeof(DataGridView).InvokeMember("DoubleBuffered",
                    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.SetProperty,
                    null, dgv, new object[] { true });
            }
            catch { }

            dgv.EnableHeadersVisualStyles = false;
            dgv.BorderStyle = BorderStyle.None;
            dgv.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            dgv.GridColor = Color.FromArgb(210, 218, 226); // Slate 300 soft

            // Header Style: Arial Narrow, non in grassetto, contrasto morbido e professionale
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(232, 236, 242);   // Soft Light Slate
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.FromArgb(51, 65, 85);       // Slate 700
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Arial Narrow", 9.75f, FontStyle.Regular);
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
            dgv.ColumnHeadersDefaultCellStyle.Padding = new Padding(3, 2, 3, 2);
            dgv.ColumnHeadersHeight = 26;

            // Row Default Style
            dgv.DefaultCellStyle.BackColor = Color.White;
            dgv.DefaultCellStyle.ForeColor = Color.FromArgb(15, 23, 42); // Slate 900
            dgv.DefaultCellStyle.Font = new Font("Arial Narrow", 9.5f, FontStyle.Regular);
            dgv.DefaultCellStyle.SelectionBackColor = Color.FromArgb(219, 234, 254); // Blue 100
            dgv.DefaultCellStyle.SelectionForeColor = Color.FromArgb(15, 23, 42); // Slate 900
            dgv.DefaultCellStyle.Padding = new Padding(2);

            // Alternating Row Style
            dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(248, 250, 252); // Slate 50
            dgv.AlternatingRowsDefaultCellStyle.ForeColor = Color.FromArgb(15, 23, 42);
            dgv.AlternatingRowsDefaultCellStyle.SelectionBackColor = Color.FromArgb(219, 234, 254);
            dgv.AlternatingRowsDefaultCellStyle.SelectionForeColor = Color.FromArgb(15, 23, 42);

            dgv.RowTemplate.Height = 24;
            dgv.BackgroundColor = Color.White;
        }

        public static void StyleButton(Button btn, Image icon = null, Color? backColor = null, Color? foreColor = null)
        {
            if (btn == null) return;
            btn.FlatStyle = FlatStyle.Flat;
            btn.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225);
            btn.FlatAppearance.BorderSize = 1;
            btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(241, 245, 249);
            btn.FlatAppearance.MouseDownBackColor = Color.FromArgb(226, 232, 240);
            btn.BackColor = backColor ?? Color.FromArgb(248, 250, 252);
            btn.ForeColor = foreColor ?? Color.FromArgb(15, 23, 42);
            btn.Font = new Font("Segoe UI", 9.0f, FontStyle.Bold);
            btn.Cursor = Cursors.Hand;

            if (icon != null)
            {
                btn.Image = icon;
                btn.ImageAlign = ContentAlignment.MiddleLeft;
                btn.TextAlign = ContentAlignment.MiddleCenter;
                btn.TextImageRelation = TextImageRelation.ImageBeforeText;
            }
        }

        public static ToolStripRenderer GetModernMenuRenderer()
        {
            return new ModernMenuRenderer();
        }

        public static void SaveFormBounds(Form form)
        {
            if (form == null) return;
            try
            {
                string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "APOffice");
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                string filePath = Path.Combine(dir, "FormLayouts.ini");

                Rectangle bounds = form.WindowState == FormWindowState.Normal ? form.Bounds : form.RestoreBounds;
                string state = form.WindowState.ToString();

                string section = form.Name;
                string value = string.Format("{0},{1},{2},{3},{4}", bounds.X, bounds.Y, bounds.Width, bounds.Height, state);

                Dictionary<string, string> dict = new Dictionary<string, string>();
                if (File.Exists(filePath))
                {
                    string[] lines = File.ReadAllLines(filePath);
                    foreach (string line in lines)
                    {
                        int idx = line.IndexOf('=');
                        if (idx > 0)
                        {
                            string k = line.Substring(0, idx).Trim();
                            string v = line.Substring(idx + 1).Trim();
                            dict[k] = v;
                        }
                    }
                }

                dict[section] = value;

                using (StreamWriter sw = new StreamWriter(filePath, false, Encoding.UTF8))
                {
                    foreach (var kvp in dict)
                    {
                        sw.WriteLine("{0}={1}", kvp.Key, kvp.Value);
                    }
                }
            }
            catch { }
        }

        public static void RestoreFormBounds(Form form)
        {
            if (form == null) return;
            try
            {
                string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "APOffice", "FormLayouts.ini");
                if (!File.Exists(filePath)) return;

                string[] lines = File.ReadAllLines(filePath);
                foreach (string line in lines)
                {
                    int idx = line.IndexOf('=');
                    if (idx > 0)
                    {
                        string k = line.Substring(0, idx).Trim();
                        if (k.Equals(form.Name, StringComparison.OrdinalIgnoreCase))
                        {
                            string v = line.Substring(idx + 1).Trim();
                            string[] parts = v.Split(',');
                            if (parts.Length >= 4)
                            {
                                int x = int.Parse(parts[0]);
                                int y = int.Parse(parts[1]);
                                int w = int.Parse(parts[2]);
                                int h = int.Parse(parts[3]);

                                Rectangle rect = new Rectangle(x, y, w, h);
                                bool isVisibleOnAnyScreen = false;
                                foreach (Screen screen in Screen.AllScreens)
                                {
                                    if (screen.WorkingArea.IntersectsWith(rect))
                                    {
                                        isVisibleOnAnyScreen = true;
                                        break;
                                    }
                                }

                                if (isVisibleOnAnyScreen)
                                {
                                    form.StartPosition = FormStartPosition.Manual;
                                    form.Location = new Point(x, y);
                                    form.Size = new Size(Math.Max(w, form.MinimumSize.Width), Math.Max(h, form.MinimumSize.Height));
                                }

                                if (parts.Length >= 5 && parts[4] == "Maximized")
                                {
                                    form.WindowState = FormWindowState.Maximized;
                                }
                            }
                            break;
                        }
                    }
                }
            }
            catch { }
        }

        public static void SaveGridColumnWidths(DataGridView dgv, string gridKey)
        {
            if (dgv == null || string.IsNullOrEmpty(gridKey)) return;
            try
            {
                string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "APOffice");
                if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);
                string filePath = Path.Combine(dir, "FormLayouts.ini");

                List<string> colData = new List<string>();
                foreach (DataGridViewColumn col in dgv.Columns)
                {
                    if (!string.IsNullOrEmpty(col.Name) && col.Visible)
                    {
                        colData.Add(string.Format("{0}:{1}", col.Name, col.Width));
                    }
                }

                if (colData.Count == 0) return;
                string value = string.Join(";", colData.ToArray());

                Dictionary<string, string> dict = new Dictionary<string, string>();
                if (File.Exists(filePath))
                {
                    string[] lines = File.ReadAllLines(filePath);
                    foreach (string line in lines)
                    {
                        int idx = line.IndexOf('=');
                        if (idx > 0)
                        {
                            string k = line.Substring(0, idx).Trim();
                            string v = line.Substring(idx + 1).Trim();
                            dict[k] = v;
                        }
                    }
                }

                dict[gridKey] = value;

                using (StreamWriter sw = new StreamWriter(filePath, false, Encoding.UTF8))
                {
                    foreach (var kvp in dict)
                    {
                        sw.WriteLine("{0}={1}", kvp.Key, kvp.Value);
                    }
                }
            }
            catch { }
        }

        public static void RestoreGridColumnWidths(DataGridView dgv, string gridKey)
        {
            if (dgv == null || string.IsNullOrEmpty(gridKey)) return;
            try
            {
                string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "APOffice", "FormLayouts.ini");
                if (!File.Exists(filePath)) return;

                string[] lines = File.ReadAllLines(filePath);
                foreach (string line in lines)
                {
                    int idx = line.IndexOf('=');
                    if (idx > 0)
                    {
                        string k = line.Substring(0, idx).Trim();
                        if (k.Equals(gridKey, StringComparison.OrdinalIgnoreCase))
                        {
                            string v = line.Substring(idx + 1).Trim();
                            string[] colPairs = v.Split(';');
                            foreach (string pair in colPairs)
                            {
                                string[] kv = pair.Split(':');
                                if (kv.Length == 2 && dgv.Columns.Contains(kv[0]))
                                {
                                    int w;
                                    if (int.TryParse(kv[1], out w) && w > 10)
                                    {
                                        dgv.Columns[kv[0]].Width = w;
                                    }
                                }
                            }
                            break;
                        }
                    }
                }
            }
            catch { }
        }

        private static void FillRoundedRectangleHelper(Graphics g, Brush brush, float x, float y, float width, float height, float radius)
        {
            using (GraphicsPath path = RoundedRect(x, y, width, height, radius))
            {
                g.FillPath(brush, path);
            }
        }

        private static void DrawRoundedRectangleHelper(Graphics g, Pen pen, float x, float y, float width, float height, float radius)
        {
            using (GraphicsPath path = RoundedRect(x, y, width, height, radius))
            {
                g.DrawPath(pen, path);
            }
        }

        private static GraphicsPath RoundedRect(float x, float y, float width, float height, float radius)
        {
            GraphicsPath path = new GraphicsPath();
            float diameter = radius * 2;
            path.AddArc(x, y, diameter, diameter, 180, 90);
            path.AddArc(x + width - diameter, y, diameter, diameter, 270, 90);
            path.AddArc(x + width - diameter, y + height - diameter, diameter, diameter, 0, 90);
            path.AddArc(x, y + height - diameter, diameter, diameter, 90, 90);
            path.CloseFigure();
            return path;
        }

        private class ModernMenuRenderer : ToolStripProfessionalRenderer
        {
            public ModernMenuRenderer() : base(new ModernColorTable()) { }

            protected override void OnRenderMenuItemBackground(ToolStripItemRenderEventArgs e)
            {
                if (e.Item.Selected)
                {
                    Rectangle rc = new Rectangle(2, 1, e.Item.Width - 4, e.Item.Height - 2);
                    using (SolidBrush b = new SolidBrush(Color.FromArgb(224, 231, 255)))
                    using (Pen p = new Pen(Color.FromArgb(165, 180, 252)))
                    {
                        e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                        FillRoundedRectangleHelper(e.Graphics, b, rc.X, rc.Y, rc.Width, rc.Height, 3);
                        DrawRoundedRectangleHelper(e.Graphics, p, rc.X, rc.Y, rc.Width, rc.Height, 3);
                    }
                }
                else
                {
                    base.OnRenderMenuItemBackground(e);
                }
            }
        }

        private class ModernColorTable : ProfessionalColorTable
        {
            public override Color MenuStripGradientBegin => Color.FromArgb(241, 245, 249);
            public override Color MenuStripGradientEnd => Color.FromArgb(226, 232, 240);
            public override Color MenuItemSelected => Color.FromArgb(224, 231, 255);
            public override Color MenuItemBorder => Color.FromArgb(199, 210, 254);
            public override Color MenuBorder => Color.FromArgb(203, 213, 225);
            public override Color ToolStripDropDownBackground => Color.White;
            public override Color ImageMarginGradientBegin => Color.FromArgb(248, 250, 252);
            public override Color ImageMarginGradientMiddle => Color.FromArgb(248, 250, 252);
            public override Color ImageMarginGradientEnd => Color.FromArgb(248, 250, 252);
            public override Color SeparatorDark => Color.FromArgb(226, 232, 240);
            public override Color SeparatorLight => Color.White;
        }
    }
}
