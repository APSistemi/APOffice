using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.IO;
using System.Threading;

namespace APOffice
{
    class clsFtp
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        public string _strFtpPath = "";
        public string _strFtpHost = "";
        public string _strFtpUser = "";
        public string _strFtpPswd = "";
        public string _strLocPath = "";
        public string _strDivTip = "";

        public string FtpTest()
        {
            string sMsg = "";

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
                request.Credentials = new NetworkCredential(_strFtpUser, _strFtpPswd);

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

            return sMsg;
        }

        public void FtpUpLoad(string strFil)
        {
            //FtpWebRequest ftpWebRequest = (FtpWebRequest)WebRequest.Create("ftp://" + _strFtpHost + _strFtpPath + Path.GetFileName(strFil));

            //ftp://www.apsistemi.net/www.apsistemi.net/ApPhoneDiv/ApPhone07/

            Console.WriteLine("aaaa");

            //string s = "ftp://" + _strFtpHost + _strFtpPath + Path.GetFileName(strFil);
            string s = "ftp://" + _strFtpHost + _strFtpPath + Path.GetFileName(strFil);

            Console.WriteLine("aaaa");

            FtpWebRequest ftpWebRequest = (FtpWebRequest)WebRequest.Create(s);
            //ftpWebRequest.Method = "STOR";
            ftpWebRequest.Method = WebRequestMethods.Ftp.UploadFile;
            
            ftpWebRequest.Credentials = (ICredentials)new NetworkCredential(_strFtpUser, _strFtpPswd);
            StreamReader streamReader = new StreamReader(strFil);
            byte[] buffer = System.IO.File.ReadAllBytes(strFil);
            streamReader.Close();
            ftpWebRequest.ContentLength = (long)buffer.Length;
            Stream requestStream = ftpWebRequest.GetRequestStream();
            requestStream.Write(buffer, 0, buffer.Length);
            requestStream.Close();
            FtpWebResponse ftpWebResponse = (FtpWebResponse)ftpWebRequest.GetResponse();
            Console.WriteLine("Upload File Complete, status {0}", (object)ftpWebResponse.StatusDescription);
            ftpWebResponse.Close();
            requestStream.Close();
            streamReader.Close();
        }

        public string FtpDownLoad(string strFil)
        {
            string sMsg = "";

            try
            {
                FileInfo fi = new FileInfo(_strLocPath + Path.GetFileName(strFil));
                FtpWebRequest ftpRequest = (FtpWebRequest)WebRequest.Create("ftp://" + _strFtpHost + _strFtpPath + Path.GetFileName(strFil));

                ftpRequest.Credentials = new NetworkCredential(_strFtpUser, _strFtpPswd);
                ftpRequest.Method = WebRequestMethods.Ftp.DownloadFile;
                ftpRequest.UsePassive = true;
                ftpRequest.EnableSsl = false;

                FtpWebResponse ftpResponse = (FtpWebResponse)ftpRequest.GetResponse();
                Stream responseStream = ftpResponse.GetResponseStream();

                byte[] buffer = new byte[1024];
                int bytesRead = responseStream.Read(buffer, 0, 1024);

                if (bytesRead > 0)
                {
                    FileStream fs = new FileStream(fi.FullName, FileMode.Create, FileAccess.Write);

                    while (bytesRead != 0)
                    {
                        fs.Write(buffer, 0, bytesRead);
                        try
                        {
                            bytesRead = responseStream.Read(buffer, 0, 1024);
                        }
                        catch (WebException e1)
                        {
                            sMsg = e1.Message;
                            break;
                        }
                    }
                    fs.Close();
                    fs.Dispose();
                }

                responseStream.Flush();
                responseStream.Close();
                responseStream.Dispose();
                ftpResponse.Close();
                ftpRequest = null;
                ftpResponse = null;
            }
            catch (WebException e)
            {
                sMsg += e.Message;
            }

            return sMsg;
        }

        public ArrayList FtpRead(ref string strMsg, ArrayList aryPre)
        {
            ArrayList aryFil = new ArrayList();
            Boolean b = true;

            try
            {
                _clsFun.ErrorLog("000", "ftp://" + _strFtpHost + _strFtpPath);

                FtpWebRequest ftpRequest = (FtpWebRequest)WebRequest.Create("ftp://" + _strFtpHost + _strFtpPath);
                //ftpRequest = (FtpWebRequest)WebRequest.Create("ftp://" + _strFtpHost + _strFtpPath);
                ftpRequest.Credentials = new NetworkCredential(_strFtpUser, _strFtpPswd);
                ftpRequest.Method = WebRequestMethods.Ftp.ListDirectory;
                ftpRequest.EnableSsl = false;
                ftpRequest.UsePassive = true;

                try
                {
                    _clsFun.ErrorLog("001", "");

                    FtpWebResponse ftpResponse = (FtpWebResponse)ftpRequest.GetResponse();
                    StreamReader sr = new StreamReader(ftpResponse.GetResponseStream(), System.Text.Encoding.ASCII);

                    while (sr.Peek() >= 0)
                    {

                        Thread.Sleep(500);

                        string sFil = sr.ReadLine();

                        if (sFil.Substring(0, 2) == "FA")
                            Console.WriteLine("aaaaa");


                        if (sFil.Length > 3)
                        {
                            b = false;

                            if (sFil.Length > 8 && sFil.Substring(0, 8).ToUpper() == "KONZMOVI")
                                Console.WriteLine("Salto");
                            else
                            {

                                foreach (string sPar in aryPre)
                                {
                                    if (sFil.Length >= sPar.Length && sFil.Substring(0, sPar.Length).ToLower() == sPar.ToLower())
                                    {
                                        b = true;
                                        break;
                                    }
                                }

                                if (b)
                                {
                                    if (File.Exists(_strLocPath + "\\" + sFil))
                                        strMsg += sFil + " già scaricato \r\n";
                                    else
                                        aryFil.Add(sFil);
                                }
                            }
                        }
                        //break;
                    }

                    _clsFun.ErrorLog("003", aryFil.Count.ToString());


                    sr.Close();
                    sr.Dispose();
                    ftpResponse.Close();
                    ftpRequest = null;
                }
                catch (WebException e)
                {
                    strMsg = e.Message;
                    ftpRequest = null;
                }

            }
            catch (WebException e)
            {
                strMsg = "Comunicazione con " + _strFtpHost + _strFtpPath + " non riuscita";
                b = false;
            }

            return aryFil;
        }

        public string FtpListDir()
        {
            FtpWebRequest ftpWebRequest = (FtpWebRequest)WebRequest.Create("ftp://" + _strFtpHost + _strFtpPath);
            ftpWebRequest.Method = WebRequestMethods.Ftp.ListDirectoryDetails;

            ftpWebRequest.Credentials = (ICredentials)new NetworkCredential(this._strFtpUser, this._strFtpPswd);

            FtpWebResponse response = (FtpWebResponse)ftpWebRequest.GetResponse();

            Stream responseStream = response.GetResponseStream();
            StreamReader reader = new StreamReader(responseStream);

            StringBuilder sbResults = new StringBuilder();

            string s = "";
            string sRig = reader.ReadLine();
            while (sRig != null)
            {
                s += sRig.Substring(35) + "|";
                //sbResults.Append(line);
                //sbResults.Append("\n");
                sRig = reader.ReadLine();
            }

            reader.Close();
            response.Close();

            return s;
        }

        public string FtpRename(string strFil, string strRen)
        {
            string sMsg = "";

            FtpWebRequest ftpRequest;

            try
            {
                if (!File.Exists(_strLocPath + strFil))
                    sMsg += strFil + " non importato \r\n";
                else
                {
                    FileInfo fi = new FileInfo(_strLocPath + strFil);
                    if (fi.Length == 0)
                        sMsg += strFil + " non importato correttamente \r\n";
                    else
                    {

                        ftpRequest = (FtpWebRequest)WebRequest.Create("ftp://" + _strFtpHost + _strFtpPath + strFil);
                        ftpRequest.Credentials = new NetworkCredential(_strFtpUser, _strFtpPswd);
                        ftpRequest.Method = WebRequestMethods.Ftp.Rename;
                        ftpRequest.UsePassive = true;
                        ftpRequest.RenameTo = strRen;
                        ftpRequest.EnableSsl = false;
                        FtpWebResponse ftpResponse = (FtpWebResponse)ftpRequest.GetResponse();
                        ftpResponse.Close();
                        ftpResponse = null;
                        ftpRequest = null;
                    }
                }
            }
            catch (WebException e)
            {
                sMsg = e.Message;
                ftpRequest = null;
            }

            return sMsg;
        }

        public string FtpDelete(string strFil)
        {
            string sMsg = "";

            FtpWebRequest ftpRequest;

            try
            {
                if (!File.Exists(_strLocPath + strFil))
                    sMsg += strFil + " non importato \r\n";
                else
                {
                    FileInfo fi = new FileInfo(_strLocPath + strFil);
                    if (fi.Length == 0)
                        sMsg += strFil + " non importato correttamente \r\n";
                    else
                    {

                        ftpRequest = (FtpWebRequest)WebRequest.Create("ftp://" + _strFtpHost + _strFtpPath + strFil);
                        ftpRequest.Credentials = new NetworkCredential(_strFtpUser, _strFtpPswd);
                        ftpRequest.Method = WebRequestMethods.Ftp.DeleteFile;
                        ftpRequest.UsePassive = true;

                        //ftpRequest.RenameTo = strRen;
                        ftpRequest.EnableSsl = false;
                        FtpWebResponse ftpResponse = (FtpWebResponse)ftpRequest.GetResponse();
                        ftpResponse.Close();
                        ftpResponse = null;
                        ftpRequest = null;
                    }
                }
            }
            catch (WebException e)
            {
                sMsg = e.Message;
                ftpRequest = null;
            }

            return sMsg;
        }

    }

}