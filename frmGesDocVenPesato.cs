using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO.Ports;
using System.Threading;

namespace APOffice
{
    public partial class frmGesDocVenPesato : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        public Boolean _bolNoVirg = false;
        public string _strLot = "";
        public string _strNegLic = "";      //Listino di cessione
        public string _strNegLir = "";      //Listino di rivendita
        public string _strArt = "";
        public string _strUmi = "";
        public string _strPxc = "";
        public string _strEan = "";
        public decimal _decPce = 0;         //Cessione 
        public decimal _decPve = 0;
        public decimal _decPes = 0;
        public decimal _decTar = 0;
        public string _strTipEti = "";

        //Etichetta
        public string _strConSql = "";
        public string _strCli = "";
        public string _strCld = "";
        public string _strInd = "";
        public string _strCpa = "";
        public string _strArd = "";
        public int _intScaGio = 0;
        public string _strBilTime = "";

        public string _strPrnLabelAuto = "S";

        public DataTable _tabArt = new DataTable();
        public DataTable _tabRow = new DataTable();
        public DataTable _tabTgr = new DataTable();

        private Color _colFocus = Color.PaleGoldenrod;
        private Color _colOld = new Color();

        public string _strOpenPorta = "";
        public SerialPort _serialPort;

        private Boolean _bolPortaAperta = false;
        private string _strIni11VenditaTouch = "";

        public frmGesDocVenPesato()
        {
            InitializeComponent();
        }

        private void frmGesScoPlu_Load(object sender, EventArgs e)
        {
            btnNVir.Enabled = _bolNoVirg;

            //lblArtPez.Visible = false;
            //lblBarPes.Visible = true;

            _strIni11VenditaTouch = _clsFun.FileIni("R", clsDefine.enuIni.Ini11VenditaTouch, "");

            if (_strCpa.Length > 0 && _strCpa.Substring(0, 1) == "S")
                nomeClienteToolStripMenuItem.Text = nomeClienteToolStripMenuItem.Text.Replace("NO", "SI");
            else
                nomeClienteToolStripMenuItem.Text = nomeClienteToolStripMenuItem.Text.Replace("SI", "NO");

            if (_strCpa.Length > 2 && _strCpa.Substring(2, 1) == "S")
                PrezzoEtichettaMenuItem1.Text = PrezzoEtichettaMenuItem1.Text.Replace("NO", "SI");
            else
                PrezzoEtichettaMenuItem1.Text = PrezzoEtichettaMenuItem1.Text.Replace("SI", "NO");

            CtrlButtons();
            FillImg();
            _colOld = txtPve.BackColor;

            if (_strOpenPorta != "S")
                FillCom();
            else if (_serialPort.IsOpen)
                _bolPortaAperta = true;

            FillDati();
        }
        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }
        private void frmGesScoPlu_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
        }
        private void Esci()
        {
            if (etichettaSIToolStripMenuItem.Text != "Etichetta SI")
                _strPrnLabelAuto = "N";
            if (chkEan128.Checked)
                _strTipEti = "EAN128";

            if (_strOpenPorta != "S")
            { 
                _clsFun.ErrorLog("Uscita", "Passo");

                try
                {
                    if (_serialPort.IsOpen)
                    {
                        _serialPort.Close();
                        _serialPort.Dispose();
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("errore su chiusura perchè gia chiusa");
                    _clsFun.ErrorLog("Uscita", "Errore");
                }
            }

            this.Close();
        }

        private void FillCom()
        {
            string s = "COM5,9600,8,0,1,10,10,0";
            s = _clsFun.FileIni("R", clsDefine.enuIni.Ini12BilCheckOut, "");
            if (s != "")
            {
                _serialPort = new SerialPort();

                string[] a = s.Split(',');

                _serialPort.PortName = Convert.ToString(a[0]);
                _serialPort.BaudRate = Convert.ToInt32(a[1]);
                _serialPort.DataBits = Convert.ToInt16(a[2]);
                _serialPort.StopBits = (StopBits)Enum.Parse(typeof(StopBits), a[4]);
                //_serialPort.Handshake = (Handshake)Enum.Parse(typeof(Handshake), a[4]);
                _serialPort.Handshake = (Handshake)Enum.Parse(typeof(Handshake), "0");
                //_serialPort.Parity = (Parity)Enum.Parse(typeof(Parity), a[7]);

                _bolPortaAperta = false;

                string[] ports = SerialPort.GetPortNames();

                //Thread.Sleep(5000);

                //MessageBox.Show("ZZZZZZZZZ");

                // Display each port name to the console.
                foreach (string port in ports)
                {
                    Console.WriteLine(port);

                    if (port == _serialPort.PortName)
                        _bolPortaAperta = true;
                }

                if (!_bolPortaAperta)
                    MessageBox.Show("Accesso alla bilancia non riuscito!", "CONTROLLO BILANCIA", MessageBoxButtons.OK, MessageBoxIcon.Error);
                else
                {

                    if (_serialPort.IsOpen)
                        MessageBox.Show("Porta già aperta");
                    else
                    {
                        Boolean bErr = true;

                        for (int i = 0; i < 10; i++)
                        {
                            try
                            {
                                //Thread.Sleep(100);
                                _serialPort.Open();

                                bErr = false;

                                _clsFun.ErrorLog("Aperto", i.ToString());

                                break;
                            }
                            catch (Exception ex)
                            {
                                //MessageBox.Show("Accesso alla bilancia non riuscito!", "CONTROLLO PORTA", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                Console.WriteLine("Error");

                                _clsFun.ErrorLog("Entrata", i.ToString() + " " + ex.Message);

                                if (_serialPort.IsOpen)
                                {
                                    bErr = false;
                                    break;
                                }
                            }
                        }
                        if (bErr)
                            MessageBox.Show("Accesso alla bilancia non riuscito!", "CONTROLLO PORTA", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            timer1.Enabled = false;
            BilInterroga();
            timer1.Enabled = true;
        }

       private void BilInterroga()
        {
            if (_bolPortaAperta && _serialPort.IsOpen)
            {
                try
                {
                    string sRig = _serialPort.ReadExisting();

                    int iTime = 100;
                    if (_clsFun.Numerico(_strBilTime))
                        iTime = Convert.ToInt32(_strBilTime);
                    Thread.Sleep(iTime);

                    string[] a = sRig.Split(new[] { "\n" }, StringSplitOptions.None);

                    for (int i = 0; i < a.Length - 1; i++)
                    {
                        //Console.Write(a[i]);

                        string ss = a[i];

                        string[] aa = ss.Split(',');

                        if(aa.Length > 2 && (aa[1] == "ST"  || aa[1] == "US"))
                        {

                            string sPes = aa[2].Trim();
                            Console.WriteLine(sPes);

                            if (_clsFun.Numerico(sPes))
                            {
                                txtPes.Text = sPes;
                                break;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }

        private void FillDati()
        {
            _strArt = "";

            if (_tabArt.Rows.Count > 0)
            {
                timer1.Enabled = true;

                _strArt = (string)_tabArt.Rows[0]["tmp_art"];
                _strUmi = (string)_tabArt.Rows[0]["tmp_umi"];
                _strPxc = Convert.ToString(_tabArt.Rows[0]["tmp_pxc"]);
                lblArd.Text = (string)_tabArt.Rows[0]["tmp_ard"];

                /* Listino di vendita pubblico per l'etichetta */
                decimal d = (decimal)_tabArt.Rows[0]["tmp_prp"];
                txtPve.Text = (d * 100).ToString("0");
                lblPve.Text = "(€ " + d.ToString("#0.00") + ")";

                /* Listino di RIvendita */
                d = (decimal)_tabArt.Rows[0]["tmp_prv"];
                txtPce.Text = (d * 100).ToString("0");
                lblPce.Text = "(€ " + d.ToString("#0.00") + ")";

                if (txtPes.Text == "" && _decPes > 0)
                {
                    if(_strUmi == "NR")
                        txtPes.Text = Convert.ToInt32(_decPes).ToString();
                    else
                        txtPes.Text = Convert.ToInt32(_decPes * 1000).ToString();
                }

                d = (decimal)_tabArt.Rows[0]["tmp_tar"];
                txtTar.Text = d.ToString("0");

                if (_strLot == "")
                    _strLot = _clsQry.ArtLotto(_strArt);

                //if (_strLot == "")
                //    _strLot = DateLotto();

                txtGsc.Text = _intScaGio.ToString();

                txtLot.Text = _strLot;

                _tabRow = new clsGenTabTmp().TabTmpRows("RowPes");
            }
            else
            { 
                SeekEcrArt();
                chkPesMan.Checked = false;
                chkPesMan.Text = "Peso da bilancia";
                timer1.Enabled = true;
            }

            if (_strUmi == "NR")
            {
                txtTar.Enabled = false;
                lblPes.Text = "Pezzi";
                lblPne.Text = "";
                chkPesMan.Checked = false;
                chkPesMan.Text = "Peso manuale";
                timer1.Enabled = false;
                //txtPes.Text = "0";
            }
            else
            {
                txtTar.Enabled = true;
                lblPes.Text = "Peso gr.";
                chkPesMan.Checked = true;
                timer1.Enabled = true;
            }

            if (!_bolPortaAperta)
            {
                txtBackColor(txtPes);
            }
            else
            {
                if(txtPes.Text == "")
                    txtBackColor(txtPes);
                else
                    txtBackColor(txtPve);
            }
        }

        private void CtrlButtons()
        {
            Image imgBtn1 = new Bitmap(@"C:\ApProject\Temp\Img\ApShop\btn03.png");

            //string s = "SELECT * FROM " + TABFON;
            //DataTable tFon = _clsFun.FillTabMdb(TABFON, s, false, _clsFun.ConMdb(_clsDef.MDBFON));

            //DataRow[] j = tFon.Select("tab_cod='009'");
            //Font fonBtn = _clsFun.FontTab(tFon, "009");

            foreach (Control c in pnlTasti.Controls)
            {
                if (c.GetType() == typeof(Button) && c.Name.Substring(0, 3) == "btn")
                {
                    Button b = (Button)c;
                    b.Visible = true;
                    b.BackColor = Color.Transparent;
                    b.ForeColor = Color.White;
                    b.BackgroundImageLayout = ImageLayout.Stretch;
                    b.FlatStyle = FlatStyle.Flat;
                    b.FlatAppearance.BorderSize = 0;
                    b.BackgroundImage = imgBtn1;
                    //b.Font = fonBtn;
                }
            }
        }

        private void FillImg()
        {
            Image imgBtn1 = new Bitmap(@"C:\ApProject\Temp\Img\ApShop\btn03.png");
            Image imgBtn2 = new Bitmap(@"C:\ApProject\Temp\Img\ApShop\btn03.png");
            Image imgCalc = new Bitmap(@"C:\ApProject\Temp\Img\ApShop\btnCal.png");

            //string s = "SELECT * FROM " + TABFON;
            //DataTable tFon = _clsFun.FillTabMdb(TABFON, s, false, _clsFun.ConMdb(_clsDef.MDBFON));

            ////btnOk.Font = _clsFun.FontTab(tFon, "005");

            foreach (ToolStripMenuItem t in menuStrip1.Items)
            {
                if (t.Text != "")
                {
                    t.BackgroundImage = imgBtn1;
                    t.BackColor = Color.Transparent;
                    t.ForeColor = Color.White;
                    t.BackgroundImageLayout = ImageLayout.Stretch;
                    //toolItem.FlatStyle = FlatStyle.Flat;
                    //toolItem.FlatAppearance.BorderSize = 0;
                    //btn.Font = new System.Drawing.Font("Microsoft Sans Serif", 28.0f, System.Drawing.FontStyle.Bold);
                    //t.Font = _clsFun.FontTab(tFon, "005");
                }
            }
        }

        private void btn_Click(object sender, EventArgs e)
        {
            lblMsg.Text = "";

            Button btn = (Button)sender;
            String s = btn.Name;

            TextBox txt = txtPve;
            if (txtTar.BackColor == _colFocus)
                txt = txtTar;
            else if (txtPes.BackColor == _colFocus)
                txt = txtPes;
            else if (txtGsc.BackColor == _colFocus)
                txt = txtGsc;
            else if (txtPce.BackColor == _colFocus)
                txt = txtPce;

            if (s.Length == 7)
            {
                if (s.Substring(3, 1) == "N" || s.Substring(3, 1) == "T")
                {
                    if (s.Substring(4) == "Vir" && _bolNoVirg)
                        s = btn.Text;
                    else if (s.Substring(4) == "Ze1" || s.Substring(4) == "Ze2")
                        s = btn.Text;
                    else
                        s = Convert.ToInt16(s.Substring(4)).ToString();

                    if (txt.Text == "0")
                        txt.Text = "";
                    txt.Text += s;
                }
                else if (s.Substring(3, 1) == "F")
                {
                    s = s.Substring(4);

                    if (s == "Pme")
                    {
                        if (txt.Text != "" && txt.Text.Substring(0, 1) == "-")
                            txt.Text += txt.Text.Substring(1);
                        else
                            txt.Text = "-" + txt.Text;
                    }
                    else if (s == "Ann")
                        txt.Text = "";
                    else if (s == "Can")
                    {
                        if (txt.Text.Length > 0)
                            txt.Text = txt.Text.Substring(0, txt.Text.Length - 1);
                    }
                    else if (s != "Inv")
                        txt.Text += btn.Text;

                    else if (s == "Inv")    // && txt.Text != "")
                    {
                        string sEan = (string)_tabArt.Rows[0]["tmp_ean"];
                        //if (sEan == "")
                        //    lblMsg.Text = "BARCODE/PLU MANCANTE";

                        if (txtPve.Text == "")
                            txtPve.Text = "0";
                        decimal dPve = Convert.ToDecimal(txtPve.Text) / 100;

                        if (txtPce.Text == "")
                            txtPce.Text = "0";
                        decimal dPce = Convert.ToDecimal(txtPce.Text) / 100;

                        if (txtTar.Text == "")
                            txtTar.Text = "0";
                        decimal dTar = Convert.ToDecimal(txtTar.Text);

                        if (!_clsFun.Numerico(txtGsc.Text, "0123456789"))
                            txtGsc.Text = "0";
                        decimal dGsc = Convert.ToDecimal(txtGsc.Text);

                        if (_strCli != "" && (dPve == 0 || dPce == 0))
                            lblMsg.Text = "PREZZI ARTICOLO MANCANTI";
                        else if(dGsc == 0)
                            MessageBox.Show("Numero giorni non corretto!", "CONTROLLO GIONI SCADENZA", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        else
                        {
                            decimal d = 0;

                            s = txtPve.Text;

                            if (s.Length > 0)
                            {
                                _decPce = dPce;
                                _decPve = dPve;
                                _strEan = sEan;

                                if (!_clsFun.Numerico(txtPes.Text.Trim()))
                                    txtPes.Text = "0";
                                _decTar = Convert.ToDecimal(txtTar.Text.Trim());
                                _decPes = (Convert.ToDecimal(txtPes.Text.Trim()) - Convert.ToDecimal(txtTar.Text.Trim())) / 1000;

                                _clsFun.ErrorLog("Controllo peso", _decPes.ToString());

                                if (_decPes > 0 || _strUmi != "KG")
                                {
                                    DataRow x = _tabRow.NewRow();
                                    x["tmp_art"] = _strArt;
                                    x["tmp_umi"] = _strUmi;
                                    x["tmp_pxc"] = _strPxc;
                                    x["tmp_lot"] = txtLot.Text;
                                    x["tmp_tar"] = _decTar;
                                    x["tmp_qkg"] = _decPes;
                                    x["tmp_pve"] = _decPce;
                                    x["tmp_prp"] = _decPve;
                                    x["tmp_gsc"] = dGsc;
                                    _tabRow.Rows.Add(x);

                                    if (chkPesSing.Checked)
                                        Esci();
                                    else
                                    {
                                        if (etichettaSIToolStripMenuItem.Text == "Etichetta SI")
                                        {
                                            //EtcSingolo();
                                            if (chkEan128.Checked)
                                                EtcSingoloEan128();
                                            else
                                                EtcSingolo();
                                        }
                                    }
                                }
                                else
                                    lblMsg.Text = "QUANTITA' MANCANTI";
                            }
                            else
                                lblMsg.Text = "PREZZI ARTICOLO MANCANTI";
                        }
                    }
                }
            }
        }

        private void txtPve_Click(object sender, EventArgs e)
        {
            txtBackColor(txtPve);
        }
        
        private void txtTar_Click(object sender, EventArgs e)
        {
            txtBackColor(txtTar);
        }

        private void txtPes_Click(object sender, EventArgs e)
        {
            txtBackColor(txtPes);
        }

        private void txtGsc_Click(object sender, EventArgs e)
        {
            txtBackColor(txtGsc);
        }

        private void txtPce_Click(object sender, EventArgs e)
        {
            txtBackColor(txtPce);
        }

        private void txtPve_TextChanged(object sender, EventArgs e)
        {
            decimal d = 0;
            if (_clsFun.Numerico(txtPve.Text))
                d = Convert.ToDecimal(txtPve.Text) / 100;

            lblPve.Text = "(€ " + d.ToString("#0.00") + ")";
        }

        private void txtPce_TextChanged(object sender, EventArgs e)
        {
            decimal d = 0;
            if (_clsFun.Numerico(txtPce.Text))
                d = Convert.ToDecimal(txtPce.Text) / 100;

            lblPce.Text = "(€ " + d.ToString("#0.00") + ")";
        }

        private void txtTar_TextChanged(object sender, EventArgs e)
        {
            PesoNetto();
        }

        private void txtPes_TextChanged(object sender, EventArgs e)
        {
            PesoNetto();
        }

        private void txtGsc_TextChanged(object sender, EventArgs e)
        {
            if (_clsFun.Numerico(txtGsc.Text, "0123456789"))
                _intScaGio = Convert.ToInt16(txtGsc.Text);
        }

        private void PesoNetto()
        {
            decimal dTar = 0;
            if (_clsFun.Numerico(txtTar.Text))
                dTar = Convert.ToDecimal(txtTar.Text) / 1000;
            decimal dPes = 0;
            if (_clsFun.Numerico(txtPes.Text))
                dPes = Convert.ToDecimal(txtPes.Text) / 1000;
            dPes -= dTar;

            if(_strUmi != "NR")
                lblPne.Text = "(Peso netto Kg " + dPes.ToString("#0.000") + ")";
        }

        private void btnSeekEcr_Click(object sender, EventArgs e)
        {
            SeekEcrArt();
        }

        private void SeekEcrArt()
        {
            frmSeekEcrArt f = new frmSeekEcrArt();
            f.ShowDialog();
            if (f._strArtCod != "")
            {
                //DataTable tArt = _clsQry.SeekArtEan(f._strArtCod, "LNE" + _strNegLir);
                DataTable tArt = _clsQry.SeekArtEan(f._strArtCod, "LNE" + _strNegLic);

                if (tArt.Rows.Count > 0)
                {
                    /* Prezzo al pubblico per l'etichetta */
                    DataTable tPrp = _clsQry.ArtPrezzo(tArt.Copy(), _strNegLir);
                    if(tPrp.Rows.Count > 0)
                        tArt.Rows[0]["tmp_prp"] = tPrp.Rows[0]["tmp_prv"];

                    _tabArt = tArt.Copy();

                    if (!DBNull.Value.Equals(_tabArt.Rows[0]["tmp_gsc"]) && _clsFun.Numerico((string)_tabArt.Rows[0]["tmp_gsc"]))
                        _intScaGio = Convert.ToInt16(_tabArt.Rows[0]["tmp_gsc"]);

                    FillDati();
                }
            }
            else
                Esci();
        }

        private void etichettaSIToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (etichettaSIToolStripMenuItem.Text == "Etichetta SI")
                etichettaSIToolStripMenuItem.Text = "Etichetta NO";
            else
                etichettaSIToolStripMenuItem.Text = "Etichetta SI";
        }

        private void nomeClienteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (nomeClienteToolStripMenuItem.Text == "Nome cliente su etichetta SI")
            {
                nomeClienteToolStripMenuItem.Text = "Nome cliente su etichetta NO";
                _strCpa = "N" + _strCpa.Substring(1);
            }
            else
            {
                nomeClienteToolStripMenuItem.Text = "Nome cliente su etichetta SI";
                _strCpa = "S" + _strCpa.Substring(1);
            }
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (PrezzoEtichettaMenuItem1.Text == "Prezzo su etichetta SI")
            {
                PrezzoEtichettaMenuItem1.Text = "Prezzo su etichetta NO";
                _strCpa = _strCpa.Substring(0,2) + "N";
            }
            else
            {
                PrezzoEtichettaMenuItem1.Text = "Prezzo su etichetta SI";
                _strCpa = _strCpa.Substring(0, 2) + "S";
            }
        }

        private void chkPesMan_Click(object sender, EventArgs e)
        {
            if (chkPesMan.Checked){
                chkPesMan.Text = "Peso manuale";
                timer1.Enabled = false;
                txtBackColor(txtPes);
            }
            else
            {
                chkPesMan.Text = "Peso da bilancia";
                timer1.Enabled = true;
                txtBackColor(txtPve);

                try
                {
                    _serialPort.Close();
                    _serialPort.Open();
                }
                catch (Exception ex)
                {
                    _clsFun.ErrorLog("Chiudi - apri", "Errore");
                }

                BilInterroga();
            }
        }

        private void chkPesSing_Click(object sender, EventArgs e)
        {
            if (chkPesSing.Checked)
            {
                chkPesSing.Text = "Pesata singola";
                chkPesSing.BackColor = _colOld;
            }
            else
            {
                chkPesSing.Text = "Pesata ripetuta";
                chkPesSing.BackColor = _colFocus;
            }

        }

        private void txtBackColor(TextBox txt)
        {
            txtPve.BackColor = _colOld;
            txtTar.BackColor = _colOld;
            txtPes.BackColor = _colOld;
            txtPce.BackColor = _colOld;
            txtGsc.BackColor = _colOld;

            txt.BackColor = _colFocus;
            txt.Select();
        }

        private void btnLot_Click(object sender, EventArgs e)
        {
            frmSeekDocLotti f = new frmSeekDocLotti();
            f._strSelect = "S";
            f.ShowDialog();
            if (f._strRes != "")
            {
                string sLot = f._strRes;

                txtLot.Text = sLot;
                _strLot = sLot;
            }
        }

        private void btnPrnEti_Click(object sender, EventArgs e)
        {
            if(chkEan128.Checked)
                EtcSingoloEan128();
            else
                EtcSingolo();
        }
        
        //private void btnPrnEt2_Click(object sender, EventArgs e)
        //{
        //    EtcSingoloEan128();
        //}

        private void EtcSingolo()
        {
            string s = "";

            decimal dTar = 0;
            if (_clsFun.Numerico(txtTar.Text))
                dTar = Convert.ToDecimal(txtTar.Text) / 1000;

            decimal dPes = 0;
            if (_clsFun.Numerico(txtPes.Text))
                dPes = Convert.ToDecimal(txtPes.Text) / 1000 - dTar;

            decimal dPrv = 0;
            if (_clsFun.Numerico(txtPve.Text))
                dPrv = Convert.ToDecimal(txtPve.Text) / 100;

            decimal dImp = 0;

            if (dPes > 0 && dPrv > 0)
                dImp = Math.Round(dPes * dPrv, 2);

            if (_strUmi != "KG")
            {
                dImp = dPrv;
                decimal dTgv = 0;
                if (!DBNull.Value.Equals(_tabArt.Rows[0]["tmp_pne"]))
                    dPes = (decimal)_tabArt.Rows[0]["tmp_pne"];

                DataRow[] j = _tabTgr.Select("tab_cod='" + (string)_tabArt.Rows[0]["tmp_tgr"] + "'");
                if (j.Length > 0 && !DBNull.Value.Equals(j[0]["tab_val"]))
                    dTgv = (decimal)j[0]["tab_val"];

                if (dTgv == 0)
                    MessageBox.Show("Tipo grammatura non corretta sull'anagrafica articolo!", "CONTROLLO ARTICOLO", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                if (dPes > 0 && dTgv > 0)
                    dPrv = dImp / dPes / dTgv;
            }

            string sPrn = "";
            string sTip = "";
            string[] a = _strIni11VenditaTouch.Split(';');
            if (a.Length > 4)
                sPrn = a[4];
            if (a.Length > 5)
                sTip = a[5];

            //MessageBox.Show(_strIni11VenditaTouch, sPrn);

            if(_strCli == "")       //Carico di magazzino
            {
                Boolean b = true;
                //string sEan = _clsQry.GenLotto2Ean(_strLot, _strArt, txtPes.Text);

                //if (sEan == "" && MessageBox.Show("Barcode non trovato, continui con la stampa dell'etichetta?", "STAMPA ETICHETTA", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) 
                //    b = false;

                if (b)
                {
                    frmGesMovIngLabel05 f = new frmGesMovIngLabel05();
                    f._strEtiTipo = sTip;
                    f._strConSql = _strConSql;
                    f._strCliCod = _strCli;
                    f._strCliDes = _strCld;
                    f._strCliInd = _strInd;
                    f._strCliPar = _strCpa;
                    f._strArtCod = _strArt;
                    //f._strEanCod = sEan;
                    f._strArtDes = lblArd.Text;
                    f._strArtUmi = _strUmi;
                    f._decMovQkg = dPes;
                    f._decMovPrv = dPrv;
                    f._decMovImp = dImp;
                    f._dayArtSca = DateTime.Today.AddDays(_intScaGio);
                    f._strLotCod = _strLot;         // txtLot.Text;
                    f.ShowDialog();
                }
            }
            else if (sPrn != "")
            {
                clsGesMovIngLabel03 cls = new clsGesMovIngLabel03();
                cls._strPrn = sPrn;
                cls._strConSql = _strConSql;
                cls._strCliDes = _strCli;
                cls._strCliInd = _strInd;
                cls._strCliPar = _strCpa;
                cls._strArtCod = _strArt;
                cls._strArtDes = lblArd.Text;
                cls._strArtUmi = _strUmi;
                cls._decMovQkg = dPes;
                cls._decMovPrv = dPrv;
                cls._decMovImp = dImp;
                cls._dayArtSca = DateTime.Today.AddDays(_intScaGio);
                cls._strLotCod = txtLot.Text;
                cls.Stampa();
            }
            else
            {
                frmGesMovIngLabel01 f = new frmGesMovIngLabel01();
                f._strEtiTipo = sTip;
                f._strConSql = _strConSql;
                f._strCliCod = _strCli;
                f._strCliDes = _strCld;
                f._strCliInd = _strInd;
                f._strCliPar = _strCpa;
                f._strArtCod = _strArt;
                f._strArtDes = lblArd.Text;
                f._strArtUmi = _strUmi;
                f._decMovQkg = dPes;
                f._decMovPrv = dPrv;
                f._decMovImp = dImp;
                f._dayArtSca = DateTime.Today.AddDays(_intScaGio);
                f._strLotCod = _strLot;         // txtLot.Text;
                f.ShowDialog();
            }
        }

        private void EtcSingoloEan128()
        {
            if (!_clsFun.Numerico(txtGsc.Text.Trim(), "0123456789") || Convert.ToInt16(txtGsc.Text.Trim()) == 0)
                MessageBox.Show("Giorni scadenza mancanti in anagrafica articolo!", "CONTROLLO ARTICOLO", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            else
            {
                string sArd = lblArd.Text;

                int iGsc = Convert.ToInt16(txtGsc.Text);
                DateTime dDsc = DateTime.Today.AddDays(iGsc);
                string sLot = _strLot;
                string sPxc = _strPxc;
                //string sEan = sPxc + " " + sLot;

                //decimal dTar = 0;
                //if (_clsFun.Numerico(txtTar.Text))
                //    dTar = Convert.ToDecimal(txtTar.Text) / 1000;

                //decimal dPes = 0;
                //if (_clsFun.Numerico(txtPes.Text))
                //    dPes = Convert.ToDecimal(txtPes.Text) / 1000 - dTar;

                //decimal dPrv = 0;
                //if (_clsFun.Numerico(txtPve.Text))
                //    dPrv = Convert.ToDecimal(txtPve.Text) / 100;

                //decimal dImp = 0;

                //if (dPes > 0 && dPrv > 0)
                //    dImp = Math.Round(dPes * dPrv, 2);

                //if (_strUmi != "KG")
                //{
                //    dImp = dPrv;
                //    decimal dTgv = 0;
                //    if (!DBNull.Value.Equals(_tabArt.Rows[0]["tmp_pne"]))
                //        dPes = (decimal)_tabArt.Rows[0]["tmp_pne"];

                //    DataRow[] j = _tabTgr.Select("tab_cod='" + (string)_tabArt.Rows[0]["tmp_tgr"] + "'");
                //    if (j.Length > 0 && !DBNull.Value.Equals(j[0]["tab_val"]))
                //        dTgv = (decimal)j[0]["tab_val"];

                //    if (dTgv == 0)
                //        MessageBox.Show("Tipo grammatura non corretta sull'anagrafica articolo!", "CONTROLLO ARTICOLO", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

                //    if (dPes > 0 && dTgv > 0)
                //        dPrv = dImp / dPes / dTgv;
                //}

                string sPrn = "";
                string sTip = "";
                string[] a = _strIni11VenditaTouch.Split(';');
                if (a.Length > 4)
                    sPrn = a[4];
                if (a.Length > 5)
                    sTip = a[5];

                ////MessageBox.Show(_strIni11VenditaTouch, sPrn);
                //if (sPrn != "")
                //{
                //    clsGesMovIngLabel03 cls = new clsGesMovIngLabel03();
                //    cls._strPrn = sPrn;
                //    cls._strConSql = _strConSql;
                //    cls._strCliDes = _strCli;
                //    cls._strCliInd = _strInd;
                //    cls._strCliPar = _strCpa;
                //    cls._strArtCod = _strArt;
                //    cls._strArtDes = lblArd.Text;
                //    cls._strArtUmi = _strUmi;
                //    cls._decMovQkg = dPes;
                //    cls._decMovPrv = dPrv;
                //    cls._decMovImp = dImp;
                //    cls._dayArtSca = DateTime.Today.AddDays(_intScaGio);
                //    cls._strLotCod = txtLot.Text;
                //    cls.Stampa();
                //}
                //else
                //{
                //    frmGesMovIngLabel01 f = new frmGesMovIngLabel01();
                //    f._strEtiTipo = sTip;
                //    f._strConSql = _strConSql;
                //    f._strCliCod = _strCli;
                //    f._strCliDes = _strCld;
                //    f._strCliInd = _strInd;
                //    f._strCliPar = _strCpa;
                //    f._strArtCod = _strArt;
                //    f._strArtDes = lblArd.Text;
                //    f._strArtUmi = _strUmi;
                //    //f._strArtUmi = _strUmi;
                //    f._decMovQkg = dPes;
                //    f._decMovPrv = dPrv;
                //    f._decMovImp = dImp;
                //    f._dayArtSca = DateTime.Today.AddDays(_intScaGio);
                //    f._strLotCod = _strLot;         // txtLot.Text;
                //    f.ShowDialog();
                //}
                //else
                //{
                frmGesMovIngLabel03 f = new frmGesMovIngLabel03();
                f._strEtiTipo = sTip;
                f._strConSql = _strConSql;
                f._strCliCod = _strCli;
                f._strCliDes = _strCld;
                f._strCliInd = _strInd;
                f._strCliPar = _strCpa;
                f._strArtCod = _strArt;
                f._strArtDes = sArd;
                f._strArtUmi = _strUmi;
                //f._strArtUmi = _strUmi;
                //f._decMovQkg = dPes;
                //f._decMovPrv = dPrv;
                //f._decMovImp = dImp;
                f._dayDaySca = dDsc;
                f._strLotCod = sLot;            // txtLot.Text;
                f._strArtPxc = sPxc;            //Pezzi per confezione
                f.ShowDialog();
                //}
            }
        }

       
  


    }
}
