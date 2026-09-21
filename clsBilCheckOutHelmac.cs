using System;
using System.IO.Ports;
using System.Threading;
using System.Text;

namespace APOffice
{
    public class clsBilCheckOutHelmac
    {
        static bool _continue;
        static SerialPort _serialPort;

        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        public string _strDes = "";
        public decimal _decImp = 0;
        public decimal _decTar = 0;
        public decimal _decPeso = 0;
        public string _strMsg = "";

        public clsBilCheckOutHelmac()
        {
        }

        public void BilCheckOut()
        {
            string name;
            string message;
            StringComparer stringComparer = StringComparer.OrdinalIgnoreCase;
            //Thread readThread = new Thread(Read);

            int _intSleep = 500;
            int _intWait = 2000;

            // Create a new SerialPort object with default settings.
            _serialPort = new SerialPort();

            string s = "COM1,57600,8,0,1,10,10,0";
            s = _clsFun.FileIni("R", clsDefine.enuIni.Ini12BilCheckOut, "");

            string[] a = s.Split(',');

            //s = "";
            //s += Convert.ToString(a[0]) + "\r\n";
            //s += Convert.ToInt32(a[1]) + "\r\n";
            //s += Convert.ToInt16(a[2]) + "\r\n";
            //s += (StopBits)Enum.Parse(typeof(StopBits), a[4]) + "\r\n";
            //s += (Handshake)Enum.Parse(typeof(Handshake), a[4]) + "\r\n";
            //s += (Parity)Enum.Parse(typeof(Parity), a[7]) + "\r\n";

            _serialPort.PortName = Convert.ToString(a[0]);
            _serialPort.BaudRate = Convert.ToInt32(a[1]);
            _serialPort.DataBits = Convert.ToInt16(a[2]);
            _serialPort.StopBits = (StopBits)Enum.Parse(typeof(StopBits), a[4]);
            //_serialPort.Handshake = (Handshake)Enum.Parse(typeof(Handshake), a[4]);
            _serialPort.Handshake = (Handshake)Enum.Parse(typeof(Handshake), "1");
            _serialPort.Parity = (Parity)Enum.Parse(typeof(Parity), a[7]);


            // Set the read/write timeouts
            //_serialPort.ReadTimeout = 500;
            //_serialPort.WriteTimeout = 500;

            _continue = false;

            string[] ports = SerialPort.GetPortNames();

            // Display each port name to the console.
            foreach (string port in ports)
            {
                Console.WriteLine(port);

                if (port == _serialPort.PortName)
                    _continue = true;
            }

            if (!_continue)
                _strMsg = "ACCESSO ALLA BILANCIA NON  RIUSCITO";
            else
            {
                Console.ReadLine();

                try
                {
                    _serialPort.Open();
                }
                catch (Exception ex)
                {
                    _strMsg = "ACCESSO ALLA BILANCIA NON  RIUSCITO";
                    return;
                }
                //catch (TimeoutException)
                //{
                //    _strMsg = "ACCESSO ALLA BILANCIA NON  RIUSCITO";
                //    return;
                //}

                _continue = true;
                //readThread.Start();

                //Console.Write("Name: ");
                //name = Console.ReadLine();

                //Console.WriteLine("Type QUIT to exit");

                var e = Encoding.GetEncoding(1250);


                char cStx = (char)0x02;
                char cEtx = (char)0x03;
                char cSoh = (char)0x01;
                char cEot = (char)0x04;
                char cAck = (char)0x06;
                char cNak = (char)0x15;
                char cCnf = (char)0x30;
                char cMsg = (char)0x49;
                char cEsc = (char)0x1B;
                char cEnq = (char)0x05;

                decimal dImp = Convert.ToDecimal(_decImp);
                decimal dTara = Convert.ToDecimal(_decTar);
                string sDes = (_strDes + new string(' ', 13)).Substring(0, 13);
                object sMsg;

                //s = e.GetString(new byte[] { Convert.ToByte(cEot) });           // 
                //s += e.GetString(new byte[] { Convert.ToByte(cStx) });          // 
                //s += "05";
                //s += e.GetString(new byte[] { Convert.ToByte(cEsc) });          // 
                //s += Convert.ToInt64(dImp * 100).ToString("00000");
                //s += e.GetString(new byte[] { Convert.ToByte(cEsc) });          // 
                //s += Convert.ToInt64(dTara).ToString("0000");                   //Tara
                //s += e.GetString(new byte[] { Convert.ToByte(cEsc) });          // 
                //s += sDes;
                //s += e.GetString(new byte[] { Convert.ToByte(cEtx) });          // 


                s = e.GetString(new byte[] { Convert.ToByte(cEot) });           // 
                s += e.GetString(new byte[] { Convert.ToByte(cStx) });          // 
                s += "05";
                s += e.GetString(new byte[] { Convert.ToByte(cEsc) });          // 
                s += Convert.ToInt64(dImp * 100).ToString("00000");
                s += e.GetString(new byte[] { Convert.ToByte(cEsc) });          // 
                s += Convert.ToInt64(dTara).ToString("0000");                   // Tara
                s += e.GetString(new byte[] { Convert.ToByte(cEsc) });          // 
                s += sDes;
                s += e.GetString(new byte[] { Convert.ToByte(cEtx) });          // 

                _serialPort.Write(s);

                Boolean b = true;

                int i = 0;
                while (_continue)
                {
                    i++;
                    try
                    {
                        //Thread.Sleep(500); // Delay 100ms
                        //Thread.Sleep(_intSleep); // Delay 100ms

                        //string message = _comPort.ReadLine();
                        s = e.GetString(new byte[] { Convert.ToByte(cEot) });           // 
                        s += e.GetString(new byte[] { Convert.ToByte(cStx) });          //
                        s += "08";
                        s += e.GetString(new byte[] { Convert.ToByte(cEtx) });          // 
                        _serialPort.Write(s);

                        Thread.Sleep(100);

                        sMsg = _serialPort.ReadExisting();
                        _serialPort.DiscardInBuffer();

                        s = Convert.ToString(sMsg);

                        if (s != "")
                        {
                            //Console.Write("ricevuto:" + sMsg);

                            if (s.Length > 6)
                            {
                                string s2 = "";
                                foreach (Char c in s)
                                {
                                    if (_clsFun.Numerico(c))
                                        s2 += c.ToString();
                                    else
                                        s2 += "|";
                                }

                                a = s2.Split('|');

                                if (a.Length > 1)
                                {
                                    s = a[2];

                                    if (s == "00")
                                    {
                                        _strMsg = "";
                                        break;
                                    }
                                    else if (s == "01")
                                        _strMsg = "general error";
                                    else if (s == "02")
                                        _strMsg = "parityerror or bufferoverflow";
                                    else if (s == "10")
                                        _strMsg = "invalid recordno";
                                    else if (s == "11")
                                        _strMsg = "invalid unitprice";
                                    else if (s == "12")
                                        _strMsg = "invalid tarevalue";
                                    else if (s == "13")
                                        _strMsg = "invalid text";
                                    else if (s == "20")
                                        _strMsg = "scale is still in motion (no standstill)";
                                    else if (s == "21")
                                        _strMsg = "scale wasn’t in motion since last operation";
                                    else if (s == "22")
                                        _strMsg = "measurement is not yet finished";
                                    else if (s == "30")
                                        _strMsg = "weight is less than minimum weight";
                                    else if (s == "31")
                                        _strMsg = "scale is less than 0";
                                    else if (s == "32")
                                        _strMsg = "scale is overcrowded";

                                    _continue = false;
                                    break;
                                }
                            }

                        }
                    }
                    catch (TimeoutException)
                    {
                        Console.WriteLine("B ");
                    }

                    if (i > 100000)
                    {
                        _continue = false;
                        _strMsg = "PESO NON LETTO A";
                        break;
                    }
                }

                if (_continue)
                {
                    //s = e.GetString(new byte[] { Convert.ToByte(cEot) });         
                    s = e.GetString(new byte[] { Convert.ToByte(cEnq) });
                    _serialPort.Write(s);

                    _clsFun.ErrorLog("bbb", "Passo");

                    while (_continue)
                    {
                        i++;
                        try
                        {
                            Thread.Sleep(100);
                            sMsg = _serialPort.ReadExisting();
                            _serialPort.DiscardInBuffer();

                            s = Convert.ToString(sMsg);

                            //if (s != "")
                            //{
                            //    Console.Write("ricevuto:" + sMsg);
                            //s = Convert.ToString(sMsg);

                            _clsFun.ErrorLog("ccc", s);

                            if (s.Length > 10)
                            {

                                string s2 = "";
                                foreach (Char c in s)
                                {
                                    if (_clsFun.Numerico(c))
                                        s2 += c.ToString();
                                    else
                                        s2 += "|";
                                }

                                a = s2.Split('|');


                                _clsFun.ErrorLog("sss", s2);


                                // _decPeso = Convert.ToDecimal(s.Substring(6, 5));

                                _decPeso = Convert.ToDecimal(a[2]);

                                break;
                            }



                            //    break;
                            //}
                        }
                        catch (TimeoutException)
                        {
                            Console.WriteLine("B ");
                        }

                        if (i > 50)
                        {
                            _continue = false;
                            _strMsg = "PESO NON LETTO B";
                            break;
                        }
                    }


                }

                //readThread.Join();
                _serialPort.Close();
            }
        }

        private SerialPort iniCom()
        {

            string s = "COM1,57600,8,0,1,10,10,0";
            string[] a = s.Split(',');

            s = "";
            s += Convert.ToString(a[0]) + "\r\n";
            s += Convert.ToInt32(a[1]) + "\r\n";
            s += Convert.ToInt16(a[2]) + "\r\n";
            s += (StopBits)Enum.Parse(typeof(StopBits), a[4]) + "\r\n";
            s += (Handshake)Enum.Parse(typeof(Handshake), a[4]) + "\r\n";
            s += (Parity)Enum.Parse(typeof(Parity), a[7]) + "\r\n";

            SerialPort ComPort = new SerialPort();
            ComPort.PortName = Convert.ToString(a[0]);
            ComPort.BaudRate = Convert.ToInt32(a[1]);
            ComPort.DataBits = Convert.ToInt16(a[2]);
            ComPort.StopBits = (StopBits)Enum.Parse(typeof(StopBits), a[4]);
            ComPort.Handshake = (Handshake)Enum.Parse(typeof(Handshake), a[4]);
            ComPort.Parity = (Parity)Enum.Parse(typeof(Parity), a[7]);

            Boolean b = ComPort.RtsEnable;

            return ComPort;
        }


        public static void Read()
        {
            while (_continue)
            {
                try
                {
                    string message = _serialPort.ReadLine();
                    Console.WriteLine(message);
                }
                catch (TimeoutException) { }
            }
        }

        // Display Port values and prompt user to enter a port.
        public static string SetPortName(string defaultPortName)
        {
            string portName;

            Console.WriteLine("Available Ports:");
            foreach (string s in SerialPort.GetPortNames())
            {
                Console.WriteLine("   {0}", s);
            }

            Console.Write("Enter COM port value (Default: {0}): ", defaultPortName);
            portName = Console.ReadLine();

            if (portName == "" || !(portName.ToLower()).StartsWith("com"))
            {
                portName = defaultPortName;
            }
            return portName;
        }
        // Display BaudRate values and prompt user to enter a value.
        public static int SetPortBaudRate(int defaultPortBaudRate)
        {
            string baudRate;

            Console.Write("Baud Rate(default:{0}): ", defaultPortBaudRate);
            baudRate = Console.ReadLine();

            if (baudRate == "")
            {
                baudRate = defaultPortBaudRate.ToString();
            }

            return int.Parse(baudRate);
        }

        // Display PortParity values and prompt user to enter a value.
        public static Parity SetPortParity(Parity defaultPortParity)
        {
            string parity;

            Console.WriteLine("Available Parity options:");
            foreach (string s in Enum.GetNames(typeof(Parity)))
            {
                Console.WriteLine("   {0}", s);
            }

            Console.Write("Enter Parity value (Default: {0}):", defaultPortParity.ToString(), true);
            parity = Console.ReadLine();

            if (parity == "")
            {
                parity = defaultPortParity.ToString();
            }

            return (Parity)Enum.Parse(typeof(Parity), parity, true);
        }
        // Display DataBits values and prompt user to enter a value.
        public static int SetPortDataBits(int defaultPortDataBits)
        {
            string dataBits;

            Console.Write("Enter DataBits value (Default: {0}): ", defaultPortDataBits);
            dataBits = Console.ReadLine();

            if (dataBits == "")
            {
                dataBits = defaultPortDataBits.ToString();
            }

            return int.Parse(dataBits.ToUpperInvariant());
        }

        // Display StopBits values and prompt user to enter a value.
        public static StopBits SetPortStopBits(StopBits defaultPortStopBits)
        {
            string stopBits;

            Console.WriteLine("Available StopBits options:");
            foreach (string s in Enum.GetNames(typeof(StopBits)))
            {
                Console.WriteLine("   {0}", s);
            }

            Console.Write("Enter StopBits value (None is not supported and \n" +
             "raises an ArgumentOutOfRangeException. \n (Default: {0}):", defaultPortStopBits.ToString());
            stopBits = Console.ReadLine();

            if (stopBits == "")
            {
                stopBits = defaultPortStopBits.ToString();
            }

            return (StopBits)Enum.Parse(typeof(StopBits), stopBits, true);
        }
        public static Handshake SetPortHandshake(Handshake defaultPortHandshake)
        {
            string handshake;

            Console.WriteLine("Available Handshake options:");
            foreach (string s in Enum.GetNames(typeof(Handshake)))
            {
                Console.WriteLine("   {0}", s);
            }

            Console.Write("End Handshake value (Default: {0}):", defaultPortHandshake.ToString());
            handshake = Console.ReadLine();

            if (handshake == "")
            {
                handshake = defaultPortHandshake.ToString();
            }

            return (Handshake)Enum.Parse(typeof(Handshake), handshake, true);
        }
    }
}