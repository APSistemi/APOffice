using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO.Ports;

namespace APOffice
{
    class clsBilCheckOutSerialPort
    {
        private Boolean _bolComAttiva = true;

        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();

        public clsBilCheckOutSerialPort()
        {
        }

        public SerialPort iniCom(ref string strMsg)
        {
            SerialPort serialPort = new SerialPort();            
            //string sMsg = "";
            string s = "COM5,9600,8,0,1,10,10,0";
            s = _clsFun.FileIni("R", clsDefine.enuIni.Ini12BilCheckOut, "");
            if (s != "")
            {
                string[] a = s.Split(',');

                serialPort.PortName = Convert.ToString(a[0]);
                serialPort.BaudRate = Convert.ToInt32(a[1]);
                serialPort.DataBits = Convert.ToInt16(a[2]);
                serialPort.StopBits = (StopBits)Enum.Parse(typeof(StopBits), a[4]);
                //_serialPort.Handshake = (Handshake)Enum.Parse(typeof(Handshake), a[4]);
                serialPort.Handshake = (Handshake)Enum.Parse(typeof(Handshake), "0");
                //serialPort.Parity = (Parity)Enum.Parse(typeof(Parity), a[7]);
                serialPort.Parity = (Parity)Enum.Parse(typeof(Parity), "0");

                //_bolPortaAperta = false;

                string[] ports = SerialPort.GetPortNames();

                // Display each port name to the console.
                foreach (string port in ports)
                {
                    Console.WriteLine(port);

                    if (port == serialPort.PortName)
                    {
                        //_bolPortaAperta = true;
                        strMsg = "OK";
                    }
                }

                //if (!_bolPortaAperta)
                if (strMsg != "OK")
                {
                    //MessageBox.Show("Accesso alla bilancia non riuscito!", "CONTROLLO BILANCIA", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    strMsg = "Accesso alla bilancia non riuscito!|CONTROLLO BILANCIA";
                }
                else
                {
                    if (serialPort.IsOpen)
                    {
                        //MessageBox.Show("Porta già aperta");
                        strMsg = "Porta già aperta";
                    }
                    else
                    {
                        Boolean bErr = true;

                        for (int i = 0; i < 10; i++)
                        {
                            try
                            {
                                //Thread.Sleep(100);
                                serialPort.Open();
                                bErr = false;
                                _clsFun.ErrorLog("Aperto", i.ToString());
                                strMsg = "OK";
                                break;
                            }
                            catch (Exception ex)
                            {
                                //MessageBox.Show("Accesso alla bilancia non riuscito!", "CONTROLLO PORTA", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                Console.WriteLine("Error");

                                _clsFun.ErrorLog("Entrata", i.ToString() + " " + ex.Message);

                                if (serialPort.IsOpen)
                                {
                                    bErr = false;
                                    break;
                                }
                            }
                        }
                        if (bErr)
                        {
                            //MessageBox.Show("Accesso alla bilancia non riuscito!", "CONTROLLO PORTA", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            strMsg = "Accesso alla bilancia non riuscito!|CONTROLLO PORTA";

                        }
                    }
                }
            }

            return serialPort;
        }

        //public Boolean PortaApri(SerialPort comPort)
        //{
        //    Boolean b = true;

        //    if (_bolComAttiva)
        //    {
        //        try
        //        {
        //            comPort.Open();
        //        }
        //        catch (Exception ex)
        //        {
        //            Console.WriteLine(ex.Message);
        //            b = false;
        //        }
        //    }

        //    if (b)
        //        _clsFun.FileLog(_clsDef.FILLOG, "PrnnSco", "Porta aperta");
        //    else
        //        _clsFun.FileLog(_clsDef.FILLOG, "PrnnSco", "Errore su apertura Porta");

        //    return b;
        //}


    }
}
