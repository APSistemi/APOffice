using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net.Mail;

class clsMail
{
    public Boolean _boolHtml = false;

    public string _strSmptCli = "";
    public int _intPorta = 0;
    public string _strMittInd = "";
    public string _strMittNome = "";
    public string _strMittPwd = "";
    public string _strDestInd = "";
    public string _strDestNome = "";
    public string _strCopiaCarbone = "";
    public string _strObj = "";
    public string _strBody = "";
    public string _strAttachment = "";

    public string InvMail()
    {
        string sMsg = "";

        MailMessage mMail = new MailMessage();
        try
        {
            mMail.From = new MailAddress(_strMittInd, _strMittNome);
            mMail.To.Add(new MailAddress(_strDestInd, _strDestNome));
            if (_strCopiaCarbone != "")
            {
                //mMail.CC.Add(new MailAddress(_strCopiaCarbone, _strCopiaCarbone));
                string[] a = _strCopiaCarbone.Split(new Char[] { ';' });

                foreach (string s in a)
                {
                    if (s != "")
                        mMail.CC.Add(new MailAddress(s, s));
                }

            }

            System.Net.NetworkCredential auth = new System.Net.NetworkCredential(_strMittInd, _strMittPwd);


            mMail.Subject = _strObj;
            mMail.Body = _strBody;

            mMail.IsBodyHtml = _boolHtml;

            if (_strAttachment != "")
            {
                string[] a = _strAttachment.Split(new Char[] { ';' });

                foreach (string s in a)
                {
                    if (s != "")
                        mMail.Attachments.Add(new Attachment(s));
                }
            }

            SmtpClient smtpClient = new SmtpClient(_strSmptCli);
            smtpClient.Credentials = auth;
            smtpClient.EnableSsl = false;

            smtpClient.Send(mMail);
            //smtpClient.Send(mailMessageHTMLWithAttach);
            //smtpClient.Send(mailMessageWithEmbeddedPicture);
        }
        catch (SmtpException smtpException)
        {
            sMsg = smtpException.Message;
        }
        catch (Exception ex)
        {
            sMsg = ex.Message;
        }

        //if (sMsg != "")
        //{

        //    try
        //    {

        //        mMail.From = new MailAddress(_strMittInd, _strMittNome);
        //        mMail.To.Add(new MailAddress("paolo.secchettin@spacrovigo.org", ""));
        //        mMail.Subject = "Errore indirizzo mail " + _strDestInd;
        //        //mailMessage.CC utilizzare se è necessario inviare in copia carbone 
        //        mMail.Body = sMsg +" "+ _strObj +" "+ _strBody;

        //        SmtpClient smtpClient = new SmtpClient("192.168.1.190");

        //        smtpClient.Send(mMail);
        //    }
        //    catch (Exception ex)
        //    {
        //        sMsg += ex.Message;
        //    }
        
        //}

        return sMsg;

    }

    public string InvioMail()
    {
        string sMsg = "";

        string stReturnText = string.Empty;
        try
        {
            if (!string.IsNullOrEmpty(_strMittInd))
            {
                //Set SmtpClient to send Email
                string stFromUserName = _strMittInd;
                string stFromPassword = _strMittPwd;
                int inPort = Convert.ToInt32(587);
                string stHost = _strSmptCli;
                bool btIsSSL = true;

                MailAddress to = new MailAddress(_strDestInd);
                MailAddress from = new MailAddress(stFromUserName);

                MailMessage objEmail = new MailMessage(from, to);
                objEmail.Subject = _strObj;
                objEmail.Body = "";
                objEmail.IsBodyHtml = true;
                objEmail.Priority = MailPriority.High;

                if (_strCopiaCarbone != "")
                {
                    //mMail.CC.Add(new MailAddress(_strCopiaCarbone, _strCopiaCarbone));
                    string[] a = _strCopiaCarbone.Split(new Char[] { ';' });

                    foreach (string s in a)
                    {
                        if (s != "")
                            objEmail.CC.Add(new MailAddress(s, s));
                    }
                }


                if (_strAttachment != "")
                {
                    string[] a = _strAttachment.Split(new Char[] { ';' });

                    foreach (string s in a)
                    {
                        if (s != "")
                            objEmail.Attachments.Add(new Attachment(s));
                    }
                }

                SmtpClient client = new SmtpClient();
                System.Net.NetworkCredential auth = new System.Net.NetworkCredential(stFromUserName, stFromPassword);
                client.Host = stHost;
                client.Port = inPort;
            //    client.UseDefaultCredentials = true;

                client.Credentials = new System.Net.NetworkCredential(stFromUserName, stFromPassword);

                client.EnableSsl = btIsSSL;
                client.Send(objEmail);
                client.Dispose();
            }
        }
        catch (Exception ex)
        {
            sMsg = ex.Message;
            Console.WriteLine(sMsg);
        }

        Console.WriteLine("OK");

        return sMsg;
    }

    public string oldInvioMailAG()
    {
        string sMsg = "";

        string stReturnText = string.Empty;
        try
        {
            if (!string.IsNullOrEmpty(_strMittInd))
            {
                //_strMittInd = "secnet64@gmail.com";
                //_strMittPwd = "Paolo345!";
                //_strDestInd = "secks@tiscali.it";


                //Set SmtpClient to send Email
                string stFromUserName = _strMittInd;
                string stFromPassword = _strMittPwd;
                int inPort = Convert.ToInt32(_intPorta);
                string stHost = _strSmptCli;
                bool btIsSSL = true;

                //MailAddress to = new MailAddress(_strDestInd);
                MailAddress to = new MailAddress(_strDestInd);
                MailAddress from = new MailAddress(stFromUserName);

                MailMessage objEmail = new MailMessage(from, to);
                objEmail.Subject = _strObj;
                objEmail.Body = "";
                objEmail.IsBodyHtml = true;
                objEmail.Priority = MailPriority.High;

                if (_strCopiaCarbone != "")
                {
                    //mMail.CC.Add(new MailAddress(_strCopiaCarbone, _strCopiaCarbone));
                    string[] a = _strCopiaCarbone.Split(new Char[] { ';' });

                    foreach (string s in a)
                    {
                        if (s != "")
                            objEmail.CC.Add(new MailAddress(s, s));
                    }
                }

                if (_strAttachment != "")
                {
                    string[] a = _strAttachment.Split(new Char[] { ';' });

                    foreach (string s in a)
                    {
                        if (s != "")
                            objEmail.Attachments.Add(new Attachment(s));
                    }
                }

                SmtpClient client = new SmtpClient();
                System.Net.NetworkCredential auth = new System.Net.NetworkCredential(stFromUserName, stFromPassword);
                client.Host = stHost;       // "sendm.cert.legalmail.it"; sendm.cert.legalmail.it 
                client.Port = inPort;
                //client.Port = 25;         //465
                client.UseDefaultCredentials = false;
                client.Credentials = auth;
                client.EnableSsl = btIsSSL;
                client.Send(objEmail);

                client.Dispose();
                objEmail.Dispose();
            }
        }
        catch (Exception ex)
        {
            sMsg = ex.Message;
            Console.WriteLine(sMsg);
        }

        Console.WriteLine("OK");

        return sMsg;
    }

    public string xxxInvioMailAG()
    {
        string sMsg = "";

        string stReturnText = string.Empty;
        try
        {
            if (!string.IsNullOrEmpty(_strMittInd))
            {
                //Set SmtpClient to send Email
                string stFromUserName = _strMittInd;
                string stFromPassword = _strMittPwd;
                int inPort = Convert.ToInt32(_intPorta);
                string stHost = _strSmptCli;
                bool btIsSSL = true;

                MailAddress to = new MailAddress(_strDestInd);
                MailAddress from = new MailAddress(stFromUserName);

                MailMessage objEmail = new MailMessage(from, to);
                objEmail.Subject = _strObj;
                objEmail.Body = "";
                objEmail.IsBodyHtml = true;
                objEmail.Priority = MailPriority.High;

                if (_strCopiaCarbone != "")
                {
                    //mMail.CC.Add(new MailAddress(_strCopiaCarbone, _strCopiaCarbone));
                    string[] a = _strCopiaCarbone.Split(new Char[] { ';' });

                    foreach (string s in a)
                    {
                        if (s != "")
                            objEmail.CC.Add(new MailAddress(s, s));
                    }
                }


                if (_strAttachment != "")
                {
                    string[] a = _strAttachment.Split(new Char[] { ';' });

                    foreach (string s in a)
                    {
                        if (s != "")
                            objEmail.Attachments.Add(new Attachment(s));
                    }
                }

                SmtpClient client = new SmtpClient();
                System.Net.NetworkCredential auth = new System.Net.NetworkCredential(stFromUserName, stFromPassword);
                client.Host = stHost;       // "sendm.cert.legalmail.it"; sendm.cert.legalmail.it 
                client.Port = inPort;
                //client.Port = 25;       //465
                client.UseDefaultCredentials = false;
                client.Credentials = auth;
                client.EnableSsl = btIsSSL;
                client.Send(objEmail);

                client.Dispose();
                objEmail.Dispose();
            }
        }
        catch (Exception ex)
        {
            sMsg = ex.Message;
            Console.WriteLine(sMsg);
        }

        Console.WriteLine("OK");

        return sMsg;
    }


    public string InvioMailAG()
    {
        string sMsg = "";

        string stReturnText = string.Empty;
        try
        {
            if (!string.IsNullOrEmpty(_strMittInd))
            {
                //Set SmtpClient to send Email
                string stFromUserName = _strMittInd;
                string stFromPassword = _strMittPwd;
                int inPort = Convert.ToInt32(_intPorta);
                string stHost = _strSmptCli;
                bool btIsSSL = true;

                MailAddress to = new MailAddress(_strDestInd);
                MailAddress from = new MailAddress(stFromUserName);

                MailMessage objEmail = new MailMessage(from, to);
                objEmail.Subject = _strObj;
                objEmail.Body = "";
                objEmail.IsBodyHtml = true;
                objEmail.Priority = MailPriority.High;

                if (_strCopiaCarbone != "")
                {
                    //mMail.CC.Add(new MailAddress(_strCopiaCarbone, _strCopiaCarbone));
                    string[] a = _strCopiaCarbone.Split(new Char[] { ';' });

                    foreach (string s in a)
                    {
                        if (s != "")
                            objEmail.CC.Add(new MailAddress(s, s));
                    }
                }


                if (_strAttachment != "")
                {
                    string[] a = _strAttachment.Split(new Char[] { ';' });

                    foreach (string s in a)
                    {
                        if (s != "")
                            objEmail.Attachments.Add(new Attachment(s));
                    }
                }

                SmtpClient client = new SmtpClient();
                System.Net.NetworkCredential auth = new System.Net.NetworkCredential(stFromUserName, stFromPassword);
                client.Host = stHost;       // "sendm.cert.legalmail.it";
                client.Port = inPort;
                client.UseDefaultCredentials = false;
                client.Credentials = auth;
                client.EnableSsl = btIsSSL;
                client.Send(objEmail);
                client.Dispose();
                objEmail.Dispose();
            }
        }
        catch (Exception ex)
        {
            sMsg = ex.Message;
            Console.WriteLine(sMsg);
        }

        Console.WriteLine("OK");

        return sMsg;
    }

    public void MailProva()
    {
        string stReturnText = string.Empty;
        try
        {
            if (!string.IsNullOrEmpty("apsistemi.srl@gmail.com"))
            {
                ////Set SmtpClient to send Email
                string stFromUserName = "apsistemi.srl@gmail.com";
                string stFromPassword = "entraentra2";
                //int inPort = Convert.ToInt32(587);
                //string stHost = "smtp.gmail.com";
                //bool btIsSSL = true;

                //string stFromUserName = "secnet64@gmail.com";
                //string stFromPassword = "secnet1964";
                int inPort = Convert.ToInt32(587);
                string stHost = "smtp.gmail.com";
                bool btIsSSL = true;

                MailAddress to = new MailAddress("secks@tiscali.it");
                MailAddress from = new MailAddress(stFromUserName);

                MailMessage objEmail = new MailMessage(from, to);
                objEmail.Subject = "Prova";
                objEmail.Body = "Prova";
                objEmail.IsBodyHtml = true;
                objEmail.Priority = MailPriority.High;


                SmtpClient client = new SmtpClient();
                System.Net.NetworkCredential auth = new System.Net.NetworkCredential(stFromUserName, stFromPassword);
                client.Host = stHost;
                client.Port = inPort;
                client.UseDefaultCredentials = false;
                client.Credentials = auth;
                client.EnableSsl = btIsSSL;
                client.Send(objEmail);

            }
        }
        catch (Exception ex)
        {
            string s = ex.Message;
            Console.WriteLine(s);
        }

        Console.WriteLine("OK");

    }


    public void InvMailAtt(string strMittInd, string strMittNome, string strDestInd, string strDestNome, string strObj, string strBody, string strCC, string strAtt)
    {

        //MailMessage cMai = new MailMessage();
        //cMai.From = "miamail@miodominio.it";
        //cMai.To = TB_YourEmail.Text;
        //cMai.BodyFormat = MailFormat.Html;
        //cMai.Subject = "Promozione";
        //cMai.Body = TB_Body.Text;
        //cMai.SmtpServer = "localhost";
        //cMai.Send(mailMsg);

        //cMai.Visible = true;

        //Mail Plain Text - Senza allegato
        MailMessage mMail = new MailMessage();
        mMail.From = new MailAddress(strMittInd, strMittNome);
        mMail.To.Add(new MailAddress(strDestInd, strDestNome));
        mMail.Subject = strObj;
        mMail.IsBodyHtml = _boolHtml;

        if (strCC != "")
        {
            MailAddress copy = new MailAddress(strCC);
            mMail.CC.Add(copy);
        }

        mMail.Body = strBody;

        if (strAtt != "")
        {
            string[] a = strAtt.Split(new Char[] { ';' });

            foreach (string s in a)
            {
                if(s != "")
                    mMail.Attachments.Add(new Attachment(s));
            }
        }
        SmtpClient smtpClient = new SmtpClient("192.168.1.180");

        try
        {
            smtpClient.Send(mMail);
            //smtpClient.Send(mailMessageHTMLWithAttach);
            //smtpClient.Send(mailMessageWithEmbeddedPicture);
        }
        catch (SmtpException smtpException)
        {
            //_strMsg = smtpException.Message;
        }
        catch (Exception ex)
        {
            //_strMsg = ex.Message;
        }
    }

}
