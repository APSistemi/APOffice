using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using System.Threading;
using System.Net;

namespace APOffice
{
    public partial class frmProva : Form
    {
        private const string TABLISVEN = "GesLisVendita";

        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        private string _strConSql = "";

        public frmProva()
        {
            InitializeComponent();
        }

        private void frmProva_Load(object sender, EventArgs e)
        {
            _strConSql = _clsFun.ConSql("");

        }

        private void button1_Click(object sender, EventArgs e)
        {
            DataRow[] j;
            string s = "";

            s = "SELECT liv_art, liv_lis, liv_dti, liv_dtf FROM " + TABLISVEN + " ORDER BY liv_art, liv_dti DESC";
            DataTable tLiv = _clsFun.FillTabSql(TABLISVEN, s, false, _strConSql);
            DataColumn[] keys = new DataColumn[4];
            keys[0] = tLiv.Columns["liv_lis"];
            keys[1] = tLiv.Columns["liv_art"];
            keys[2] = tLiv.Columns["liv_dti"];
            keys[3] = tLiv.Columns["liv_dtf"];
            tLiv.PrimaryKey = keys;

            DataTable t = tLiv.Clone();
            keys = new DataColumn[4];
            keys[0] = t.Columns["liv_lis"];
            keys[1] = t.Columns["liv_art"];
            keys[2] = t.Columns["liv_dti"];
            keys[3] = t.Columns["liv_dtf"];
            t.PrimaryKey = keys;

            foreach(DataRow y in tLiv.Rows)
            {
                t.ImportRow(y);
            }


            dgv1.DataSource = t;

        }

        private void button2_Click(object sender, EventArgs e)
        {

            // Get a fresh copy of the sample PDF file
            string filename = "C:\\Temp\\Pippo.pdf";
            //File.Copy(Path.Combine("../../../../../PDFs/", filename),
            //  Path.Combine(Directory.GetCurrentDirectory(), filename), true);

            // Create the output document
            PdfDocument outputDocument = new PdfDocument();

            // Show single pages
            // (Note: one page contains two pages from the source document)
            outputDocument.PageLayout = PdfPageLayout.SinglePage;

            XFont font = new XFont("Verdana", 8, XFontStyle.Bold);
            XStringFormat format = new XStringFormat();
            format.Alignment = XStringAlignment.Center;
            format.LineAlignment = XLineAlignment.Far;
            XGraphics gfx;
            XRect box;

            // Open the external document as XPdfForm object
            XPdfForm form = XPdfForm.FromFile(filename);

            for (int idx = 0; idx < form.PageCount; idx += 2)
            {
                // Add a new page to the output document
                PdfPage page = outputDocument.AddPage();
                page.Orientation = PdfSharp.PageOrientation.Landscape;

                double width = page.Width;
                double height = page.Height;

                int rotate = page.Elements.GetInteger("/Rotate");

                gfx = XGraphics.FromPdfPage(page);

                // Set page number (which is one-based)
                form.PageNumber = idx + 1;

                box = new XRect(0, 0, width / 2, height);
                // Draw the page identified by the page number like an image
                gfx.DrawImage(form, box);

                // Write document file name and page number on each page
                box.Inflate(0, -10);
                gfx.DrawString(String.Format("- {1} -", filename, idx + 1),
                  font, XBrushes.Red, box, format);

                if (idx + 1 < form.PageCount)
                {
                    // Set page number (which is one-based)
                    form.PageNumber = idx + 2;

                    box = new XRect(width / 2, 0, width / 2, height);
                    // Draw the page identified by the page number like an image
                    gfx.DrawImage(form, box);

                    // Write document file name and page number on each page
                    box.Inflate(0, -10);
                    gfx.DrawString(String.Format("- {1} -", filename, idx + 2),
                      font, XBrushes.Red, box, format);
                }
            }

            // Save the document...
            filename = "TwoPagesOnOne_tempfile.pdf";
            outputDocument.Save(filename);
            // ...and start a viewer.
            Process.Start(filename);

        }

        private void button3_Click(object sender, EventArgs e)
        {
            string filename1 = "C:\\Temp\\Pippo1.pdf";
            string filename2 = "C:\\Temp\\Pippo1.pdf";


            // Open the input files
            PdfDocument inputDocument1 = PdfReader.Open(filename1, PdfDocumentOpenMode.Import);
            PdfDocument inputDocument2 = PdfReader.Open(filename2, PdfDocumentOpenMode.Import);

            // Create the output document
            PdfDocument outputDocument = new PdfDocument();

            // Show consecutive pages facing. Requires Acrobat 5 or higher.
            outputDocument.PageLayout = PdfPageLayout.TwoColumnLeft;

            XFont font = new XFont("Verdana", 10, XFontStyle.Bold);
            XStringFormat format = new XStringFormat();
            format.Alignment = XStringAlignment.Center;
            format.LineAlignment = XLineAlignment.Far;
            XGraphics gfx;
            XRect box;
            int count = Math.Max(inputDocument1.PageCount, inputDocument2.PageCount);
            for (int idx = 0; idx < count; idx++)
            {
                // Get page from 1st document
                PdfPage page1 = inputDocument1.PageCount > idx ?
                  inputDocument1.Pages[idx] : new PdfPage();

                // Get page from 2nd document
                PdfPage page2 = inputDocument2.PageCount > idx ?
                  inputDocument2.Pages[idx] : new PdfPage();

                // Add both pages to the output document
                page1 = outputDocument.AddPage(page1);
                page2 = outputDocument.AddPage(page2);

                //// Write document file name and page number on each page
                //gfx = XGraphics.FromPdfPage(page1);
                //box = page1.MediaBox.ToXRect();
                //box.Inflate(0, -10);
                //gfx.DrawString(String.Format("{0} • {1}", filename1, idx + 1), font, XBrushes.Red, box, format);

                //gfx = XGraphics.FromPdfPage(page2);
                //box = page2.MediaBox.ToXRect();
                //box.Inflate(0, -10);
                //gfx.DrawString(String.Format("{0} • {1}", filename2, idx + 1), font, XBrushes.Red, box, format);
            }

            // Save the document...
            const string filename = "CompareDocument1_tempfile.pdf";
            outputDocument.Save(filename);
            Process.Start(filename);

        }

        private void button4_Click(object sender, EventArgs e)
        {
            string sMsg = "";

            string _strFtpHost = "93.62.101.246/";
            string _strFtpPath = "";

            try
            {
                Console.WriteLine("aaaa");

                string s = "";

                //s = "ftp://www.apsistemi.net/www.apsistemi.net/ApPhoneDiv/ApPhone07/";

                s = "ftp://" + _strFtpHost + _strFtpPath;  //OK
                //string s = "ftp://" + _strFtpPath;

                // Get the object used to communicate with the server.
                FtpWebRequest request = (FtpWebRequest)WebRequest.Create(s);
                request.Method = WebRequestMethods.Ftp.ListDirectoryDetails;

                // This example assumes the FTP site uses anonymous logon.
                request.Credentials = new NetworkCredential("N8892", "Jess1c@99");

                FtpWebResponse response = (FtpWebResponse)request.GetResponse();

                Stream responseStream = response.GetResponseStream();
                StreamReader reader = new StreamReader(responseStream);
                Console.WriteLine(reader.ReadToEnd());

                Console.WriteLine("Directory List Complete, status {0}", response.StatusDescription);

                reader.Close();
                response.Close();
            }
            catch (WebException ex)
            {
                sMsg = ex.Message;
            }


        }

        private void button5_Click(object sender, EventArgs e)
        {

            string reqUrl = " https://dev.servizimultimediali.net/test/api/VerificaErogazioneSN.php?numeroSerialeCarta=6060003434452007&PINCeliachia=12345&importo=10";
            string accToken = "";

            var httpWebRequestQR = (HttpWebRequest)WebRequest.Create(reqUrl);
            httpWebRequestQR.ContentType = "application/json";
            httpWebRequestQR.Method = "GET";
            //httpWebRequestQR.Headers.Add("Authorization", "Bearer " + accToken);
            //httpWebRequestQR.Headers.Add("X-Signature", signatureWc);
            //httpWebRequestQR.Headers.Add("X-Nonce-Str", nonce);
            //httpWebRequestQR.Headers.Add("X-Timestamp", timestampStr);

            var httpResponseQR = (HttpWebResponse)httpWebRequestQR.GetResponse();
            using (var streamReader = new StreamReader(httpResponseQR.GetResponseStream()))
            {
                var resultQR = streamReader.ReadToEnd();
                string jsonStringsign = resultQR;

                string s = resultQR;

                Console.WriteLine("xxx");

                //Newtonsoft.Json.JsonTextReader
                //JsonTextReader reader = new JsonTextReader(jsonStringsign);
            }
        }

    }
}
