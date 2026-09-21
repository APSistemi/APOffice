using System;
using System.Collections;
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

namespace APOffice
{
    public partial class frmGesOrdFornitori : Form
    {
        clsDefine _clsDef = new clsDefine();
        clsFuncs _clsFun = new clsFuncs();
        clsQuery _clsQry = new clsQuery();

        private string _strConSql = "";

        public string _strOftYea = "";
        public string _strOftFor = "";
        public string _strOftNum = "";
        public string _strOftStd = "";

        private const string TABGESORT = "GesOrdForTestate";
        private const string TABGESORF = "GesOrdForRighe";
        private const string TABTABFIM = "TabForImport";
        private const string TABANAEAN = "AnaBarcode";

        public frmGesOrdFornitori()
        {
            InitializeComponent();
            new clsGesGraph().SetGraph(this, 0);
        }

        private void frmGesOrdFornitori_Load(object sender, EventArgs e)
        {
            ApplyModernUi();
            _strConSql = _clsFun.ConSql("");

            if(_clsFun.ParGet(clsDefine.enuParametri.ParOrdForDividi, _strConSql).Trim() == "")
                importazioneDaTerminalinoToolStripMenuItem.Enabled = false;

            lblOftYea.Text = _strOftYea;
            lblOftNum.Text = _strOftNum;
            SetDgv1();
            FillTab();
            btnInvio.Text = "Invio a " + cmbOftFor.Text.Trim();
            FillDati();

            clsUiIcons.RestoreFormBounds(this);
            clsUiIcons.RestoreGridColumnWidths(dgv1, "frmGesOrdFornitori_dgv1");
        }

        private void frmGesOrdFornitori_FormClosing(object sender, FormClosingEventArgs e)
        {
            clsUiIcons.SaveFormBounds(this);
            clsUiIcons.SaveGridColumnWidths(dgv1, "frmGesOrdFornitori_dgv1");
        }

        private void ApplyModernUi()
        {
            try
            {
                if (menuStrip1 != null)
                {
                    menuStrip1.Renderer = clsUiIcons.GetModernMenuRenderer();
                    if (esciToolStripMenuItem != null)
                        esciToolStripMenuItem.Image = clsUiIcons.GetIcon("exit", 16);
                    if (importazioneDaTerminalinoToolStripMenuItem != null)
                        importazioneDaTerminalinoToolStripMenuItem.Image = clsUiIcons.GetIcon("barcode", 16);
                    if (stampaToolStripMenuItem != null)
                        stampaToolStripMenuItem.Image = clsUiIcons.GetIcon("print", 16);
                    if (tuttoToolStripMenuItem != null)
                        tuttoToolStripMenuItem.Image = clsUiIcons.GetIcon("list", 16);
                    if (erroriToolStripMenuItem != null)
                        erroriToolStripMenuItem.Image = clsUiIcons.GetIcon("warning", 16);
                }

                this.BackColor = Color.FromArgb(243, 244, 246);

                if (btnInvio != null)
                {
                    clsUiIcons.StyleStatButton(btnInvio, btnInvio.Text, "send", 16,
                        Color.FromArgb(240, 253, 250), Color.FromArgb(153, 246, 228),
                        Color.FromArgb(45, 212, 191), Color.FromArgb(19, 78, 74), Color.FromArgb(13, 148, 136));
                }

                if (dgv1 != null)
                {
                    clsUiIcons.StyleDataGridView(dgv1);
                }
            }
            catch (Exception ex)
            {
                _clsFun.ErrorLog("frmGesOrdFornitori.ApplyModernUi", ex.Message);
            }
        }

        private void frmGesOrdFornitori_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
                Esci();
        }
        private void esciToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Esci();
        }
        private void Esci()
        {
            Boolean b = Salva();
            if (!b)
            {
                if (MessageBox.Show("Abbandono delle modifiche, confermi?", this.Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                      == DialogResult.Yes)
                    this.Close();
            }
            else this.Close();
        }

        private void btnInvio_Click(object sender, EventArgs e)
        {
            Salva();

            string s = "SELECT * FROM TabForImport WHERE tab_cod='" + _strOftFor + "'";
            DataTable t = _clsFun.FillTabSql(TABTABFIM, s, false, _strConSql);

            if (t.Rows.Count == 0)
                MessageBox.Show("Fornitore non previsto per invio ordini!");
            else if (cmbOftSta.SelectedValue.ToString() != _clsDef.DIVDAD)
                MessageBox.Show("Stato dell'ordine non è 'da inviare'!");
            else  if (DBNull.Value.Equals(t.Rows[0]["tab_tip"]) || ((string)t.Rows[0]["tab_tip"]).Trim() == "")
                MessageBox.Show("Codice identificativo tipodivulgazione non previsto in tabella divulgazione fornitori!");
            else if (DBNull.Value.Equals(t.Rows[0]["tab_neg"]) || ((string)t.Rows[0]["tab_neg"]).Trim() == "")
                MessageBox.Show("Codice del cliente non definito in tabella divulgazione fornitori!");
            else if ((DBNull.Value.Equals(t.Rows[0]["tab_uso"]) || ((string)t.Rows[0]["tab_uso"]).Trim() == "") && (DBNull.Value.Equals(t.Rows[0]["tab_mai"]) || ((string)t.Rows[0]["tab_mai"]).Trim() == ""))
                MessageBox.Show("Credenziali FTP non complete in tabella divulgazione fornitori/ordine!");
            else if ((DBNull.Value.Equals(t.Rows[0]["tab_pwo"]) || ((string)t.Rows[0]["tab_pwo"]).Trim() == "") && (DBNull.Value.Equals(t.Rows[0]["tab_mai"]) || ((string)t.Rows[0]["tab_mai"]).Trim() == ""))
                MessageBox.Show("Credenziali FTP non complete in tabella divulgazione fornitori/ordine!");
            else
            {
                DataView v = new DataView(((DataView)dgv1.DataSource).ToTable(), "orf_art <>'' AND orf_qta > 0", "", DataViewRowState.CurrentRows);
                if (v.Count == 0)
                    MessageBox.Show("Non ci sono righe ordine da inviare!");
                else
                {
                    DataTable tOrd = v.ToTable();
                    string sFil = OrdInvio(((string)t.Rows[0]["tab_tip"]).Trim(), ((string)t.Rows[0]["tab_neg"]).Trim(), ((string)t.Rows[0]["tab_ord"]).Trim(), tOrd);

                    if (sFil != "")
                    {
                        if ((string)t.Rows[0]["tab_tip"] == _clsDef.DIVEUROSPESA)
                            Console.WriteLine("aaaaaaa");
                        else if ((string)t.Rows[0]["tab_tip"] == _clsDef.DIVAPMATCH)
                        {
                            Invio2ApMatch((string)t.Rows[0]["tab_pth"], sFil);
                            cmbOftSta.SelectedValue = _clsDef.DIVDIV;
                            MessageBox.Show("ORDINE INVIATO", "CONTROLLO INVIO ORDINE", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                        }
                        else if ((string)t.Rows[0]["tab_tip"] == _clsDef.DIVDPIU)
                        {
                            //Mail
                            /**/
                            //string sParMail = _clsFun.ParGet(clsDefine.enuParametri.Par015IpMail, _strConSql);

                            string sMai = (string)t.Rows[0]["tab_mai"];

                            string[] a = sMai.Split(',');

                            if (a.Length < 4)
                                MessageBox.Show("Parametri per invio mail mancanti in tabella fornitori!", "INVIO ORDINE VIA MAIL", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            else
                            {
                                string sSmptCli = a[0].Trim();
                                string sMitUsr = a[1].Trim();
                                string sMitPwd = a[2].Trim();
                                string sMaiDest = a[3].Trim();

                                clsMail _clsMail = new clsMail();
                                _clsMail._strSmptCli = sSmptCli;
                                _clsMail._strMittInd = sMitUsr;
                                _clsMail._strMittNome = sMitUsr;
                                _clsMail._strMittPwd = sMitPwd;
                                _clsMail._strDestInd = sMaiDest;
                                _clsMail._strDestNome = sMaiDest;
                                _clsMail._strCopiaCarbone = "";
                                //_clsMail._strCopiaCarbone += "psecchettin@tiscali.it";

                                //_clsMail._strObj = "Ordine O110109.010 - 13/03/2012 - 14:16";
                                _clsMail._strObj = "Ordine " + Path.GetFileName(sFil) + " - " + DateTime.Now.ToString("dd/MM/yyyy - HH:mm");
                                _clsMail._strBody = "";
                                _clsMail._strAttachment = sFil;

                                s = _clsMail.InvioMail();

                                if (s != "")
                                    MessageBox.Show(s, "ERRORE SU INVIO ORDINE VIA MAIL" , MessageBoxButtons.OK, MessageBoxIcon.Error);
                                else
                                {
                                    cmbOftSta.SelectedValue = _clsDef.DIVDIV;
                                    MessageBox.Show("ORDINE INVIATO", "CONTROLLO INVIO ORDINE VIA MAIL", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                }
                            }
                        }
                        else if ((string)t.Rows[0]["tab_tip"] == _clsDef.DIVUNICOM)
                        {
                            //Mail
                            /**/
                            //string sParMail = _clsFun.ParGet(clsDefine.enuParametri.Par015IpMail, _strConSql);

                            string sRep = "";

                            if(v.Count > 0 && (string)v[0]["orf_art"] != "")
                            {
                                s = "SELECT ";
                                s += "AnaArticoli.art_cod, TabReparti.tab_des ";
                                s += "FROM AnaArticoli LEFT OUTER JOIN TabReparti ON AnaArticoli.art_rep = TabReparti.tab_cod ";
                                s += "WHERE (AnaArticoli.art_cod = '" + (string)v[0]["orf_art"] + "')";
                                DataTable tt = _clsFun.FillTabSql("AnaArticoli", s, true, _strConSql);
                                if (tt.Rows.Count > 0 && !DBNull.Value.Equals(tt.Rows[0]["tab_des"]))
                                    sRep = (string)tt.Rows[0]["tab_des"];
                            }

                            string[] a;
                            string sNegCod = _clsFun.FileIni("R", clsDefine.enuIni.Ini09CodiceAzienda, "");
                            if (sNegCod == "")
                                sNegCod = "001";
                            DataTable tCnf = _clsQry.ConfAzienda(sNegCod);
                            if (tCnf.Rows.Count == 0)
                                MessageBox.Show("Tabella configurazione azienda non configurata!", "INVIO ORDINE VIA MAIL", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            else
                            {
                                string sAzi = ((string)tCnf.Rows[0]["cnf_rag"]).Trim();
                                string sUniNeg = "";

                                s = ((string)t.Rows[0]["tab_neg"]).Trim();
                                //s = _clsFun.FileIni("R", clsDefine.enuIni.Ini09CodiceAzienda, "");
                                if (s.IndexOf(',') > 0)
                                {
                                    a = s.Split(',');
                                    int i = Convert.ToInt16(sNegCod) - 1;
                                    if (i <= a.Length)
                                        sUniNeg = a[i];
                                }

                                string sMai = (string)t.Rows[0]["tab_mai"];

                                a = sMai.Split(',');

                                if (a.Length < 4)
                                    MessageBox.Show("Parametri per invio mail mancanti in tabella fornitori!", "INVIO ORDINE VIA MAIL", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                else
                                {
                                    string sSmptCli = a[0].Trim();
                                    string sMitUsr = a[1].Trim();
                                    string sMitPwd = a[2].Trim();
                                    string sMaiDest = a[3].Trim();

                                    string sMaiCC = "";
                                    if (a.Length > 4)
                                        sMaiCC = a[4].Trim();

                                    clsMail _clsMail = new clsMail();
                                    _clsMail._strSmptCli = sSmptCli;
                                    _clsMail._strMittInd = sMitUsr;
                                    _clsMail._strMittNome = sMitUsr;
                                    _clsMail._strMittPwd = sMitPwd;         //???
                                    _clsMail._strDestInd = sMaiDest;
                                    _clsMail._strDestNome = sMaiDest;
                                    _clsMail._strCopiaCarbone = sMaiCC;

                                    //_clsMail._strCopiaCarbone += "psecchettin@tiscali.it";

                                    //_clsMail._strObj = "Ordine O110109.010 - 13/03/2012 - 14:16";
                                    _clsMail._strObj = "Ordine " + sAzi +" "+ sRep + " neg. " + sUniNeg;
                                    _clsMail._strBody = "";
                                    _clsMail._strAttachment = sFil;

                                    s = _clsMail.InvioMail();

                                    if (s != "")
                                        MessageBox.Show(s, "ERRORE SU INVIO ORDINE VIA MAIL", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    else
                                    {
                                        cmbOftSta.SelectedValue = _clsDef.DIVDIV;
                                        MessageBox.Show("ORDINE INVIATO", "CONTROLLO INVIO ORDINE VIA MAIL", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                    }
                                }
                            }
                        }

                        else
                        {
                            s = InvioFtp(t, sFil);

                            if (s == "")
                            {
                                cmbOftSta.SelectedValue = _clsDef.DIVDIV;
                                MessageBox.Show("Ordine inviato!");
                            }
                            else
                            {
                                _clsFun.ErrorLog(s, "Ordine a fornitore");
                                MessageBox.Show(s, "ERRORE SU INVIO ORDINE");
                            }
                        }
                    }
                }
           }

        }

        private void SetDgv1()
        {
            dgv1.AutoGenerateColumns = false;
            //dgv1.VirtualMode = true;
            //dgv1.Dock = DockStyle.Fill;
            dgv1.AllowUserToAddRows = false;
            dgv1.ReadOnly = false;
            dgv1.AllowUserToDeleteRows = false;
            //dgv1.DisplayedRowCount() = true;

            DataGridViewTextBoxColumn cTbc;
            //DataGridViewCheckBoxColumn cCbc;
            //DataGridViewComboBoxColumn cCmb;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "orf_nri";
            cTbc.Name = "Riga";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "orf_arf";
            cTbc.Name = "Cod. fornitore";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "orf_art";
            cTbc.Name = "Articolo";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "orf_ard";
            cTbc.Name = "Descrizione";
            cTbc.Width = 200;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            cTbc.MaxInputLength = 50;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "orf_umi";
            cTbc.Name = "UM";
            cTbc.Width = 30;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            cTbc.MaxInputLength = 50;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "orf_pxc";
            cTbc.Name = "PxC";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "orf_qta";
            cTbc.Name = "Q.tà";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = false;
            dgv1.Columns.Add(cTbc);
            dgv1.Columns[cTbc.Name].DefaultCellStyle.BackColor = Color.GreenYellow;

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "orf_cos";
            cTbc.Name = "Costo";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "orf_prv";
            cTbc.Name = "Prezzo vendita";
            cTbc.Width = 50;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "OrfImp";
            cTbc.Name = "Importo";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            cTbc.DefaultCellStyle.Format = "###,##0.00";
            cTbc.ReadOnly = true;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "orf_msg";
            cTbc.Name = "Messaggi";
            cTbc.Width = 150;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.MaxInputLength = 50;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "orf_arf";
            cTbc.Name = "Art. fornitore";
            cTbc.Width = 60;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "orf_ean";
            cTbc.Name = "Barcode";
            cTbc.Width = 70;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = true;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

            cTbc = new DataGridViewTextBoxColumn();
            cTbc.DataPropertyName = "OrfMdy";
            cTbc.Name = "Modificato";
            cTbc.Width = 0;
            cTbc.ValueType = typeof(string);
            cTbc.ReadOnly = false;
            //cTbc.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv1.Columns.Add(cTbc);

        }

        private void FillTab()
        {
            DataRow x;
            string p = "";
            string s = "";
            DataTable t = new DataTable();

            p = "TabStatoDivulgazioni";
            s = "SELECT * FROM " + p + " WHERE tab_var=1";
            t = _clsFun.FillTabSql(p, s, false, _strConSql);
            //x = t.NewRow();
            //x["tab_cod"] = "";
            //x["tab_des"] = "  Non definito";
            //t.Rows.InsertAt(x, 0);
            cmbOftSta.DataSource = t;
            cmbOftSta.DisplayMember = "tab_des";
            cmbOftSta.ValueMember = "tab_cod";
            cmbOftSta.SelectedValue = "0";

            p = "AnaFornitori";
            s = "SELECT for_cod, for_des FROM AnaFornitori";
            t = _clsFun.FillTabSql(p, s, false, _strConSql);
            //x = t.NewRow();
            //x["tab_cod"] = "";
            //x["tab_des"] = "  Non definito";
            //t.Rows.InsertAt(x, 0);
            cmbOftFor.DataSource = t;
            cmbOftFor.DisplayMember = "for_des";
            cmbOftFor.ValueMember = "for_cod";
            cmbOftFor.SelectedValue = _strOftFor;
        }

        private void FillDati()
        {
            string s = "";
            s = "SELECT * ";
            s += "FROM GesOrdForTestate WHERE ";
            s += "oft_yea='" + lblOftYea.Text + "' AND ";
            s += "oft_num='" + lblOftNum.Text + "' ";

            DataTable t = _clsFun.FillTabSql("OFT", s, false, _strConSql);
            if (t.Rows.Count > 0)
            {
                cmbOftSta.SelectedValue = t.Rows[0]["oft_sta"];
            }

            s = "SELECT ";
            s += "orf_yea, ";
            s += "orf_num, ";
            s += "orf_nri, ";
            s += "orf_day, ";
            s += "orf_art, ";
            s += "orf_arf, ";
            s += "orf_msg, ";
            s += "orf_cxp, ";
            s += "orf_pxc, ";
            s += "orf_ean, ";
            s += "orf_umi, ";
            s += "orf_qta, ";
            s += "orf_cos, ";
            s += "orf_prv, ";
            s += "CASE WHEN orf_ard IS NULL OR orf_ard = '' THEN AnaArticoli.art_des ELSE orf_ard END AS orf_ard ";
            s += "FROM GesOrdForRighe ";
            s += "LEFT OUTER JOIN AnaArticoli ON GesOrdForRighe.orf_art = AnaArticoli.art_cod ";
            s += "WHERE ";
            s += "orf_yea='" + lblOftYea.Text + "' AND orf_num='" + lblOftNum.Text + "'";

            t = _clsFun.FillTabSql(TABGESORF, s, false, _strConSql);
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "OrfMdy",
                Caption = "Mdy",
                MaxLength = 1,
                ReadOnly = false,
                DefaultValue = (String)""
            });
            t.Columns.Add(new DataColumn()
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "OrfImp",
                Caption = "Imp",
                ReadOnly = false,
                DefaultValue = (Decimal)0
            });
            
            foreach(DataRow y in t.Rows)
            {
                if(DBNull.Value.Equals(y["orf_qta"]))
                    y["orf_qta"] = 0;
                if(DBNull.Value.Equals(y["orf_cos"]))
                    y["orf_cos"] = 0;
                if(DBNull.Value.Equals(y["orf_pxc"]))
                    y["orf_pxc"] = 1;

                y["OrfImp"] = (decimal)y["orf_qta"] * (decimal)y["orf_pxc"] * (decimal)y["orf_cos"];
            }

            DataView v = new DataView(t, "", "", DataViewRowState.CurrentRows);

            dgv1.DataSource = v;
            Totali();
            NewRiga(false);
        }

        private void NewRiga(Boolean bolCtrl)
        {
            Boolean b = true;

            DataTable t = ((DataView)dgv1.DataSource).Table;

            DataView v = new DataView(t, "orf_art=''", "orf_art", DataViewRowState.CurrentRows);

            if (v.Count > 0)
                b = false;

            if (b)
            {
                DataRow x = t.NewRow();
                x["orf_nri"] = "xxx";
                //x["orf_for"] = _strOftFor;
                x["orf_yea"] = _strOftYea;
                x["orf_num"] = _strOftNum;
                x["orf_day"] = DateTime.Today;
                x["orf_art"] = "";
                x["orf_ard"] = "";
                x["orf_qta"] = 1;
                x["orf_cos"] = 0;
                x["orf_prv"] = 0;
                t.Rows.Add(x);

                CurrentCell("RIGNEW");
            }
        }

        private void txtSeek_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Return)
            {
                string s = txtSeek.Text.Trim();

                if (s != "" && dgv1.RowCount > 0)
                {
                    Boolean b = false;
                    int iCur = dgv1.CurrentCell.RowIndex;
                    int i = 0;

                    foreach (DataGridViewRow row in dgv1.Rows)
                    {
                        i++;

                        if (i > iCur + 1)
                        {
                            if (row.Cells["Descrizione"].Value.ToString().ToLower().Contains(s.ToLower()))
                            {
                                dgv1.Rows[row.Index].Selected = true;
                                dgv1.FirstDisplayedScrollingRowIndex = row.Index;
                                dgv1.CurrentCell = dgv1.Rows[row.Index].Cells[0];
                                b = true;
                                break;
                            }
                        }
                    }
                    if (!b)
                        dgv1.CurrentCell = dgv1.Rows[0].Cells[0];
                }
            }
        }

        private void dgv1_CurrentCellChanged(object sender, EventArgs e)
        {
            Console.WriteLine("aaaa");
            //CurrentCell("RIGDES");
        }

        private int CurrentCell(string strTip)
        {
            int iPos = 3;
            if (strTip == "RIGQTA")
                iPos = 4;
            else if (strTip == "RIGDES")
                iPos = 3;
            int i = dgv1.Rows.Count - 1;
            if (i >= 0)
            {
                try
                {
                    dgv1.BeginEdit(true);
                    //dgv1.Rows[i].Selected = true;
                    dgv1.Rows[i].Cells[iPos].Selected = true;
                    dgv1.CurrentCell = dgv1[iPos, i];
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            return i;
        }

        private void dgv1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == 2 || e.ColumnIndex == 3)
            {

                string sArt = dgv1.Rows[e.RowIndex].Cells["Articolo"].Value.ToString();

                if (e.ColumnIndex == 2 || sArt.Trim() == "")
                {

                    frmSeekArt f = new frmSeekArt();
                    f.ShowDialog();
                    if (f._tabArt != null && f._tabArt.Rows.Count > 0)
                    {
                        DataTable t = f._tabArt;
                        if (t.Rows.Count > 0)
                        {
                            CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                            if (cm.Position >= 0)
                            {
                                DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                                DataRow x = r.Row;

                                sArt = (string)t.Rows[0]["tmp_art"];

                                //t = _clsQry.ArtSeek(sArt, "SEEK");

                                t = _clsQry.ArtSeek(sArt, "FOR" + cmbOftFor.SelectedValue.ToString());

                                FillOrfRow(t, x, true);
                            }
                        }
                    }
                }
                else
                {
                    frmAnaArticolo f = new frmAnaArticolo();
                    f._strArtCod = sArt;
                    f.ShowDialog();
                }
            }
        }

        private void dgv1_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            TextBox txb = e.Control as TextBox;
            string s = "";

            if (txb != null)
            {
                txb.PreviewKeyDown += (S, E) =>
                {
                    if (E.KeyCode == Keys.Enter && s == "")
                    {
                        if (dgv1.CurrentCell.OwningColumn.Name.Equals("Descrizione"))
                        {
                            Console.WriteLine("aaaaaaa");
                            if (txb.Text != "")
                            {
                                DataTable t = _clsQry.ArtSeek(txb.Text, "SEEK");
                                if (t != null && t.Rows.Count > 0)
                                {
                                    CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
                                    if (cm.Position >= 0)
                                    {
                                        DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                                        DataRow x = r.Row;
                                        FillOrfRow(t, x, true);
                                        s = "Fatto";
                                    }
                                }
                            }
                        }
                        s = "Fatto";
                    }
                    return;
                };
            }
        }

        private void dgv1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
            if (cm.Position >= 0)
            {
                DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
                DataRow x = r.Row;
                x["OrfImp"] = (decimal)x["orf_cos"] * (decimal)x["orf_qta"];

                x["OrfMdy"] = "S";
                CtrlNumRiga(x);
                cm.EndCurrentEdit();
                Totali();
                NewRiga(true);
            }
        }

        private void dgv1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dgv1.Columns[e.ColumnIndex].Name == "Ann." && dgv1.CurrentCell is DataGridViewCheckBoxCell)
            {
                dgv1.Rows[e.RowIndex].Cells["Modificato"].Value = "S";
            }
        }

        private void dgv1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dgv1.CurrentCell is DataGridViewCheckBoxCell)
            {
                dgv1.CommitEdit(DataGridViewDataErrorContexts.Commit);
            }
        }
        
        private void dgv1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (dgv1.Columns[e.ColumnIndex].Name == "Q.tà")
                e.CellStyle.BackColor = Color.GreenYellow;
        }

        private void FillOrfRow(DataTable tabTmp, DataRow x, Boolean bolNew)
        {
            DataView v = new DataView(((DataView)dgv1.DataSource).ToTable(), "orf_art='" + (string)tabTmp.Rows[0]["tmp_art"] + "'", "", DataViewRowState.CurrentRows);
            if (v.Count > 0)
            {
                if ((string)v[0]["orf_nri"] != "")
                {
                    string sArt = (string)tabTmp.Rows[0]["tmp_art"];

                    MessageBox.Show("Articolo già presente " + tabTmp.Rows[0]["tmp_art"] + " " + tabTmp.Rows[0]["tmp_ard"] + "!");
                    if (!bolNew)
                        x.Delete();
                    else
                    {
                        //x["orf_arf"] = (string)tabTmp.Rows[0]["tmp_arf"];

                        int i = 0;
                        Boolean b = false;

                        x["orf_arf"] = (string)tabTmp.Rows[0]["tmp_arf"];

                        foreach (DataGridViewRow row in dgv1.Rows)
                        {
                            i++;

                            if (row.Cells["Articolo"].Value.ToString() == sArt)
                            {
                                dgv1.Rows[row.Index].Selected = true;
                                dgv1.FirstDisplayedScrollingRowIndex = row.Index;
                                dgv1.CurrentCell = dgv1.Rows[row.Index].Cells[0];
                                b = true;
                                break;
                            }
                        }
                        if (!b)
                            dgv1.CurrentCell = dgv1.Rows[0].Cells[0];

                    }

                }
            }
            else
            {
                
                x["orf_art"] = "";
                x["orf_ard"] = "";
                x["orf_umi"] = "";
                x["orf_cos"] = 0;
                x["orf_prv"] = 0;
                x["OrfImp"] = 0;

                if (!DBNull.Value.Equals(tabTmp.Rows[0]["tmp_art"]))
                    x["orf_art"] = tabTmp.Rows[0]["tmp_art"];

                if (!DBNull.Value.Equals(tabTmp.Rows[0]["tmp_ard"]))
                    x["orf_ard"] = tabTmp.Rows[0]["tmp_ard"];

                if (!DBNull.Value.Equals(tabTmp.Rows[0]["tmp_umi"]))
                    x["orf_umi"] = tabTmp.Rows[0]["tmp_umi"];

                if (!DBNull.Value.Equals(tabTmp.Rows[0]["tmp_cos"]))
                    x["orf_cos"] = tabTmp.Rows[0]["tmp_cos"];

                if (!DBNull.Value.Equals(tabTmp.Rows[0]["tmp_prv"]))
                    x["orf_prv"] = tabTmp.Rows[0]["tmp_prv"];

                if ((decimal)x["orf_cos"] > 0 && (decimal)x["orf_qta"] > 0)
                    x["OrfImp"] = (decimal)x["orf_cos"] * (decimal)x["orf_qta"];

                x["OrfMdy"] = "S";
                //x["orf_msg"] = "";

                DataTable tArf = _clsQry.AssFor(_strOftFor, (string)x["orf_art"]);
                if (tArf.Rows.Count == 0)
                    x["orf_msg"] = "Non in assortimento";
                else
                {

                    x["orf_arf"] = "";
                    x["orf_pxc"] = 0;
                    x["orf_cxp"] = 0;
                    x["orf_cos"] = 0;
                    x["orf_ean"] = "";

                    if (!DBNull.Value.Equals(tArf.Rows[0]["lia_arf"]))
                        x["orf_arf"] = ((string)tArf.Rows[0]["lia_arf"]).Trim();

                    if (!DBNull.Value.Equals(tArf.Rows[0]["lia_pxc"]))
                        x["orf_pxc"] = tArf.Rows[0]["lia_pxc"];

                    if (!DBNull.Value.Equals(tArf.Rows[0]["lia_cxp"]))
                        x["orf_cxp"] = tArf.Rows[0]["lia_cxp"];

                    if (!DBNull.Value.Equals(tArf.Rows[0]["lia_cos"]))
                        x["orf_cos"] = tArf.Rows[0]["lia_cos"];

                    if (!DBNull.Value.Equals(tArf.Rows[0]["ean_ean"]))
                        x["orf_ean"] = tArf.Rows[0]["ean_ean"];
                }

                if (bolNew)
                {
                    int i = CurrentCell("");

                    if (i >= 0 && dgv1.Rows[i].Cells[0].Value.ToString() != "" && dgv1.Rows.Count - 1 == i)
                    {
                        NewRiga(true);
                    }
                }
            }
        }

        private void CtrlNumRiga(DataRow rowMov)
        {
            if ((string)rowMov["orf_nri"] == "xxx" && ((string)rowMov["orf_art"]) != "")
            {
                DataTable t = ((DataView)dgv1.DataSource).ToTable();
                DataView v = new DataView(t, "orf_nri<>'xxx'", "orf_nri DESC", DataViewRowState.CurrentRows);
                int i = 0;

                if (v.Count > 0)
                    i = Convert.ToInt32(v[0]["orf_nri"]);
                rowMov["orf_nri"] = (i + 1).ToString("000");
            }
        }

        private void Totali()
        {
            decimal dCos = 0;
            int iRig = 0;
            int iCol = 0;

            DataTable t = ((DataView)dgv1.DataSource).Table;

            foreach (DataRow y in t.Rows)
            {
                if (!DBNull.Value.Equals(y["orf_art"]) && (string)y["orf_art"] != "" && (decimal)y["orf_qta"] > 0)
                {
                    iRig++;
                    iCol += Convert.ToInt32(y["orf_qta"]);

                    decimal dQta = 0;
                    decimal dPxc = 0;
                    decimal dPco = 0;

                    if (!DBNull.Value.Equals(y["orf_qta"]))
                        dQta = Convert.ToDecimal(y["orf_qta"]);
                    if (!DBNull.Value.Equals(y["orf_pxc"]))
                        dPxc = Convert.ToDecimal(y["orf_pxc"]);
                    if (!DBNull.Value.Equals(y["orf_cos"]))
                        dPco = Convert.ToDecimal(y["orf_cos"]);

                    dCos += dQta * dPxc * dPco;
                }
            }

            lblOrfCol.Text = iCol.ToString("#####0");
            lblOrfRig.Text = iRig.ToString("#####0");
            lblOrfTot.Text = dCos.ToString("#####0.00");
        }

        private Boolean OrfCtrl()
        {
            Boolean b = true;

            DataView v = new DataView(((DataView)dgv1.DataSource).ToTable(), "orf_art <>'' AND orf_qta > 0", "", DataViewRowState.CurrentRows);

            if (v.Count == 0)
            { 
                b = false;
                MessageBox.Show("Non ci sono righe valide!");
            }

            return b;
        }

        private Boolean Salva()
        {
            Boolean b = OrfCtrl();
            string s = "";

            if (b)
            {
                DataTable t = new DataTable();
                ArrayList aWhe = new ArrayList();
                ArrayList aExl = new ArrayList();

                if (lblOftNum.Text == _clsDef.CODNEW)
                    lblOftNum.Text = _clsFun.NewNum(DateTime.Today.Year.ToString("0000"), clsDefine.enuNumeratori.NumGesOrdFor, 5, _strConSql);

                s = "SELECT * FROM " + TABGESORT + " WHERE oft_yea='" + lblOftYea.Text + "' AND  oft_num='" + lblOftNum.Text + "'";
                t = _clsFun.FillTabSql(TABGESORT, s, false, _strConSql);

                DataRow y = t.NewRow();
                y["oft_yea"] = lblOftYea.Text;
                y["oft_num"] = lblOftNum.Text;
                y["oft_for"] = _strOftFor;
                y["oft_day"] = DateTime.Today;
                y["oft_sta"] = cmbOftSta.SelectedValue;
                _strOftStd = cmbOftSta.Text;

                if (t.Rows.Count == 0)
                {
                    s = _clsFun.SqlInsertRow(TABGESORT, t, y);
                }
                else
                {
                    y["oft_day"] = t.Rows[0]["oft_day"];

                    aWhe = new ArrayList();
                    aExl = new ArrayList();
                    aWhe.Add("oft_yea");
                    aWhe.Add("oft_num");
                    s = _clsFun.SqlUpdRow(TABGESORT, t, t.Rows[0], y, aWhe, aExl);
                }
                if (s != "")
                {
                    _clsFun.SqlWrite(s, _strConSql);
                    _clsFun.FileLog(TABGESORT, lblOftYea.Text + "-" + lblOftNum.Text, s);
                }

                s = "SELECT * FROM GesOrdForRighe WHERE ";
                s += "orf_yea = '" + lblOftYea.Text + "' AND ";
                s += "orf_num = '" + lblOftNum.Text + "' ";
                s += "ORDER BY orf_nri";

                t = _clsFun.FillTabSql(TABGESORF, s, false, _strConSql);
                DataRow[] j;

                aWhe = new ArrayList();
                aWhe.Add("orf_yea");
                aWhe.Add("orf_num");
                aWhe.Add("orf_nri");

                aExl = new ArrayList();

                foreach (DataRow x in (((DataView)dgv1.DataSource).ToTable()).Rows)
                {
                    if ((string)x["OrfMdy"] == "S" && ((string)x["orf_art"]).Trim() != "")
                    {
                        b = true;

                        s = "orf_nri='" + x["orf_nri"] + "'";

                        j = t.Select(s);
                        if (j.Length > 0)
                        {

                            s = _clsFun.SqlUpdRow(TABGESORF, t, j[0], x, aWhe, aExl);
                        }
                        else
                        {
                            x["orf_yea"] = lblOftYea.Text;
                            x["orf_num"] = lblOftNum.Text;

                            s = _clsFun.SqlInsertRow(TABGESORF, t, x);

                        }
                        if (s != "")
                        {
                            _clsFun.SqlWrite(s, _strConSql);
                            _clsFun.FileLog("GesOrdFornitori", (string)x["orf_art"], s);
                        }
                    }
                }
            }

            return b;
        }

        private string OrdInvio(string strTip, string strNeg, string strFil, DataTable tabOrd)
        {
            string s = "";
            string sFil = "";
            Boolean b = true;

            if (strTip.Trim() == _clsDef.DIVTERRON)
            {
                sFil = _clsDef.PATHORD + strFil + DateTime.Now.ToString("yyyyMMddHHmmss");

                StreamWriter sw = new StreamWriter(sFil, true);

                string sRig = "00";
                sRig += DateTime.Now.ToString("HHmmss");
                sRig += strNeg.PadLeft(6, Convert.ToChar("0"));
                sRig += DateTime.Now.ToString("yyyyMMddHHmmss");
                sRig += "C";
                sw.Write(sRig + _clsDef.CRLF);

                foreach (DataRow y in tabOrd.Rows)
                {
                    if (DBNull.Value.Equals(y["orf_ean"]))
                        y["orf_ean"] = "";
                    if (DBNull.Value.Equals(y["orf_arf"]))
                        y["orf_arf"] = "";

                    b = true;
                    if (("" + y["orf_art"]).Trim() == "")
                        b = false;
                    if (((string)y["orf_arf"]).Trim() == "" && ((string)y["orf_ean"]).Trim() == "")
                        b = false;

                    if(b)
                    {
                        sRig = "1";
                        sRig += "000000";

                        if ((string)y["orf_arf"] == "082181")
                            Console.WriteLine("aaaa");

                        //s = ((string)y["orf_arf"]).PadLeft(6, Convert.ToChar("0"));
                        if (DBNull.Value.Equals(y["orf_arf"]) || ((string)y["orf_art"]).Trim() == "")
                            s = new string(' ', 6);
                        else
                        {
                            s = ((string)y["orf_arf"]).PadLeft(6, Convert.ToChar("0"));
                            s = s.Substring(0, 6);
                        }
                        sRig += s;

                        //Il barcode ha priorità sull'articolo del fornitore
                        if (s.Trim() != "" && s != "000000")
                            s = new string(' ', 13);
                        else
                            s = ((string)y["orf_ean"]).PadLeft(13, Convert.ToChar("0"));
                        sRig += s;

                        s = Convert.ToInt16(y["orf_qta"]).ToString("00000");
                        sRig += s;

                        sRig += ".00";

                        sw.Write(sRig + _clsDef.CRLF);
                    }

                }

                ((TextWriter)sw).Flush();
                sw.Close();
                sw.Dispose();
            }
            if (strTip.Trim() == _clsDef.DIVBRENDO)
            {
                string sNeg = strNeg.PadLeft(6, Convert.ToChar('0'));

                sFil = _clsDef.PATHORD + strFil + sNeg + DateTime.Now.ToString("yyyyMMddHHmmss") + ".DAT";

                StreamWriter sw = new StreamWriter(sFil, true);

                string sRig = "o";
                sRig += "0" + sNeg + "0" + sNeg;
                sRig += DateTime.Now.ToString("ddMMyy");
                sRig += new string(' ', 12);
                sw.Write(sRig + _clsDef.CRLF);

                foreach (DataRow y in tabOrd.Rows)
                {
                    b = true;
                    if (("" + y["orf_art"]).Trim() == "")
                        b = false;
                    if (("" + y["orf_arf"]).Trim() == "")
                        b = false;

                    if (b)
                    {
                        //s = "200000" + (string)y["orf_arf"];
                        //s = s + new clsCtrlCodici().FindMod10Digit(s);
                        //sRig = s;

                        try
                        {
                            s = "200000" + (string)y["orf_arf"];
                            s = s + new clsCtrlCodici().FindMod10Digit(s);
                            sRig = s;


                            sRig += new string('0', 6);

                            s = (Convert.ToInt16(y["orf_qta"]) * 1000).ToString("0000000");
                            sRig += s;

                            sRig += new string('0', 7);
                        }
                        catch (Exception ex)
                        {
                            sRig = "";

                            MessageBox.Show("Articolo " + (string)y["orf_art"] + " (" + (string)y["orf_arf"] + ")" + (string)y["orf_ard"], "ERRORI SU ARTICOLO", MessageBoxButtons.OK, MessageBoxIcon.Stop);

                        }

                        sw.Write(sRig + _clsDef.CRLF);
                    }

                }

                ((TextWriter)sw).Flush();
                sw.Close();
                sw.Dispose();

            }
            if (strTip.Trim() == _clsDef.DIVEUROSPESA)
            {
                string sNum = lblOftNum.Text.Substring(2,3);
                string sNeg = strNeg;

                sFil = _clsDef.PATHORD + strFil + sNeg + "." + sNum;

                if(File.Exists((sFil)))
                    File.Delete(sFil);

                StreamWriter sw = new StreamWriter(sFil, true);

                //string sRig = "100" + sNeg;
                string sRig = "100IN";
                sRig += DateTime.Now.ToString("MMdd");
                sRig += "000004 " + "";
                sw.Write(sRig + _clsDef.CRLF);

                sRig = "2";
                sRig += sNum;
                sRig += DateTime.Now.ToString("HHmmss");
                sRig += "01";
                sRig += "ZZZP";
                sw.Write(sRig + _clsDef.CRLF);

                sRig = "3";
                sRig += sNeg.PadLeft(6,Convert.ToChar("0"));
                sRig += "000000";
                sRig += " ";
                sRig += "  ";
                sw.Write(sRig + _clsDef.CRLF);

                sRig = "400000000       ";
                sw.Write(sRig + _clsDef.CRLF);

                foreach (DataRow y in tabOrd.Rows)
                {
                    b = true;
                    if (("" + y["orf_art"]).Trim() == "")
                        b = false;
                    //if (("" + y["orf_arf"]).Trim() == "" && ((string)y["orf_ean"]).Trim() == "")
                    //    b = false;

                    if (b)
                    {
                        //s = "200000" + (string)y["orf_arf"];
                        //s = s + new clsCtrlCodici().FindMod10Digit(s);
                        sRig = "5";

                        s = ((string)y["orf_arf"]).PadLeft(9,Convert.ToChar("0"));
                        sRig += s;

                        s = (Convert.ToInt16(y["orf_qta"])).ToString("0000");
                        sRig += s;

                        sRig += "00";

                        sw.Write(sRig + _clsDef.CRLF);
                    }

                }

                ((TextWriter)sw).Flush();
                sw.Close();
                sw.Dispose();

            }
            if (strTip.Trim() == _clsDef.DIVDPIU)
            {
                string sNum = lblOftNum.Text.Substring(2, 3);
                string sNeg = strNeg;

                sFil = _clsDef.PATHORD + strFil + DateTime.Now.ToString("yyMMdd") + "." + sNum;

                string sFat = "";

                string[] a = strNeg.Split(',');

                sNeg = a[0];
                sFat = a[1];  //017979

                if (File.Exists((sFil)))
                    File.Delete(sFil);

                StreamWriter sw = new StreamWriter(sFil, true);

                string sRigIni = "ORD";                             //tipo = new char[3] -> fisso 'ORD'
                sRigIni += "G" + sNeg;                            //negozio = new char[4]; "G950"
                sRigIni += sFat; // "017979";                                //cliente = new char[6] -> "016645" codice fatturazione Maxi Di
                sRigIni += "001";                                   //reparto = new char[3] -> fisso '001'
                sRigIni += "001";                                   //numeroTerminale = new char[3] -> fisso '001'
                sRigIni += "000000000001";                          //serialeTerminale = new char[30] -> fisso '000000000001
                sRigIni += "".PadLeft(18,Convert.ToChar(' '));
                sRigIni += DateTime.Now.ToString("yyyyMMddHHmmss"); //annoEsportazione = new char[4];meseEsportazione = new char[2];giornoEsportazione = new char[2];
                sRigIni += sNum.PadLeft(9, Convert.ToChar('0'));    //numeroEsportazione = new char[9] -> viene incrementato ad ogni file generato

                foreach (DataRow y in tabOrd.Rows)
                {
                    if (DBNull.Value.Equals(y["orf_msg"]))
                        y["orf_msg"] = "";

                    if (!DBNull.Value.Equals(y["orf_arf"]) && (string)y["orf_arf"] != "" && ((string)y["orf_msg"]).Trim() == "")
                    {
                        b = true;
                        if (("" + y["orf_art"]).Trim() == "")
                            b = false;

                        if (b)
                        {
                            string sRig = sRigIni;               
                            sRig += ((string)y["orf_nri"]).PadLeft(12, Convert.ToChar('0'));    //numeroRiga = new char[12]

                            s = ((string)y["orf_arf"]).PadRight(13, Convert.ToChar(" "));        //articolo = new char[13];  codice articolo Dial lungo 7 caratteri allineato a sx
                            sRig += s;

                            s = ((string)y["orf_ean"]).PadLeft(13, Convert.ToChar(" "));        //barcode = new char[13];
                            sRig += s;

                            s = (Convert.ToInt16(y["orf_qta"])).ToString("00000");               //quantita = new char[5];
                            sRig += s;

                            sRig += "000000000";                                                //prezzoRilevato = new char[9]; valorizzato a 0
                            //sRig += "00000000";                                             //anno mese giorno di rilevazione
                            //sRig += "   ";                                                  //Causale
                            //sRig += " ";                                                    //Stato di validità

                            //sRig += "00000000";                                             //anno mese giorno di rilevazione

                            sRig += DateTime.Now.ToString("yyyyMMdd");                      //anno mese giorno di rilevazione
                            sRig += "000";
                            //causale = new char[3]; non valorizzato
                            sRig += "               ";                                    //statoValidita = new char[1]; non valorizzato
                            //annoChiamata = new char[4]; non valorizzato
                            //meseChiamata = new char[2]; non valorizzato
                            //giornoChiamata = new char[2];non valorizzato
                            //oraChiamata = new char[6]; non valorizzato

                            sw.Write(sRig + _clsDef.CRLF);
                        }
                    }
                }

                ((TextWriter)sw).Flush();
                sw.Close();
                sw.Dispose();
            }
            if (strTip.Trim() == _clsDef.DIVMIGROSS)
            {
                string sNum = lblOftNum.Text.Substring(2, 3);
                string sNeg = strNeg;

                sFil = _clsDef.PATHORD + strFil + strNeg + "_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".TXT";

                if (File.Exists((sFil)))
                    File.Delete(sFil);

                StreamWriter sw = new StreamWriter(sFil, true);

                string sRig = "1";                               //fisso '1'
                sRig += strNeg;                                  // 2 – 7	= codice cliente
                sRig += "  ";                                    // 8 – 9	= “TO” se ordine su promozione
                sRig += "     ";                                 //10 – 14	= progressivo della promozione

                sw.Write(sRig + _clsDef.CRLF);

                foreach (DataRow y in tabOrd.Rows)
                {
                    if (DBNull.Value.Equals(y["orf_msg"]))
                        y["orf_msg"] = "";

                    if (!DBNull.Value.Equals(y["orf_arf"]) && (string)y["orf_arf"] != "" && ((string)y["orf_msg"]).Trim() == "")
                    {
                        b = true;
                        if (("" + y["orf_art"]).Trim() == "")
                            b = false;

                        if (b)
                        {
                            sRig = "3";      //1		= “3”
                            //sRig += ((string)y["orf_nri"]).PadLeft(12, Convert.ToChar('0'));    //numeroRiga = new char[12]

                            s = ((string)y["orf_arf"]).PadRight(9, Convert.ToChar(" "));        //2 – 8	= codice articolo (7) + . 9-10	= differenziatore (2)
                            sRig += s;

                            //s = ((string)y["orf_ean"]).PadLeft(13, Convert.ToChar(" "));        //barcode = new char[13];
                            //sRig += s;

                            s = (Convert.ToInt16(y["orf_qta"])).ToString("0000");               //11 – 14	= quantità (allineato a dx con zeri davanti)
                            sRig += s;

                            //sRig += "000000000";                                                //prezzoRilevato = new char[9]; valorizzato a 0 + 
                            ////sRig += "00000000";                                             //anno mese giorno di rilevazione
                            ////sRig += "   ";                                                  //Causale
                            ////sRig += " ";                                                    //Stato di validità

                            ////sRig += "00000000";                                             //anno mese giorno di rilevazione

                            //sRig += DateTime.Now.ToString("yyyyMMdd");                      //anno mese giorno di rilevazione
                            //sRig += "000";
                            ////causale = new char[3]; non valorizzato
                            //sRig += "               ";                                    //statoValidita = new char[1]; non valorizzato
                            ////annoChiamata = new char[4]; non valorizzato
                            ////meseChiamata = new char[2]; non valorizzato
                            ////giornoChiamata = new char[2];non valorizzato
                            ////oraChiamata = new char[6]; non valorizzato

                            sw.Write(sRig + _clsDef.CRLF);
                        }
                    }
                }

                ((TextWriter)sw).Flush();
                sw.Close();
                sw.Dispose();

            }
            if (strTip.Trim() == _clsDef.DIVUNICOM)
            {
                string sNum = lblOftNum.Text.Substring(2, 3);
                string sNeg = strNeg;

                s = _clsFun.FileIni("R", clsDefine.enuIni.Ini09CodiceAzienda, "");
                if(sNeg.IndexOf(',') > 0)
                {
                    string[] a = strNeg.Split(',');
                    int i = Convert.ToInt16(s) - 1;
                    if(i <= a.Length)
                        sNeg = a[i];
                    else
                    {
                        MessageBox.Show("Codice negozio non definito in tabella fornitori", "CONTROLLO FORNITORE", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        b = false;
                    }
                }

                if (b)
                {
                    sFil = _clsDef.PATHORD + strFil + "_" + DateTime.Now.ToString("yyyyMMdd") + "_" + DateTime.Now.ToString("HHmmss") + ".TXT";

                    if (File.Exists((sFil)))
                        File.Delete(sFil);

                    StreamWriter sw = new StreamWriter(sFil, true);

                    string sRig = "";                               //fisso '1'

                    foreach (DataRow y in tabOrd.Rows)
                    {
                        if (DBNull.Value.Equals(y["orf_msg"]))
                            y["orf_msg"] = "";
                        
                        if (!DBNull.Value.Equals(y["orf_arf"]) && (string)y["orf_arf"] != "" && ((string)y["orf_msg"]).Trim() == "")
                        {
                            b = true;
                            if (("" + y["orf_art"]).Trim() == "")
                                b = false;

                            if (b)
                            {
                                sRig = sNeg;

                                s = ((string)y["orf_arf"]).PadRight(7, Convert.ToChar("0"));
                                sRig += s;

                                //s = ((string)y["orf_ean"]).PadLeft(13, Convert.ToChar(" "));        //barcode = new char[13];
                                //sRig += s;

                                s = (Convert.ToInt16(y["orf_qta"])).ToString("000000");
                                sRig += s;

                                sw.Write(sRig + _clsDef.CRLF);
                            }
                        }
                    }

                    ((TextWriter)sw).Flush();
                    sw.Close();
                    sw.Dispose();
                }

            }
            if (strTip.Trim() == _clsDef.DIVAPMATCH)
            {
                DataRow[] j;

                s = "";

                s = "SELECT ";
                s += "AnaBarcode.ean_ean, ";
                s += "AnaBarcode.ean_art ";
                //s += "AnaArticoli.art_des ";
                s += "FROM AnaBarcode INNER JOIN AnaArticoli ON AnaBarcode.ean_art = AnaArticoli.art_cod ";
                DataTable tEan = _clsFun.FillTabSql(TABANAEAN, s, false, _strConSql);
                DataColumn[] keys = new DataColumn[2];
                keys[0] = tEan.Columns["ean_art"];
                keys[1] = tEan.Columns["ean_ean"];
                tEan.PrimaryKey = keys;

                string sNum = lblOftNum.Text.Substring(2, 3);
                string sNeg = strNeg;

                //sFil = _clsDef.PATHORD + strFil + sNeg + "." + sNum;

                sFil = _clsDef.PATHORD + strFil;

                //if (File.Exists((sFil)))
                //    File.Delete(sFil);

                StreamWriter sw = new StreamWriter(sFil, true);

                //string sRig = "100" + sNeg;
                string sRig = "100IN";
                //sRig += DateTime.Now.ToString("MMdd");
                //sRig += "000004 " + "";
                //sw.Write(sRig + _clsDef.CRLF);

                //sRig = "2";
                //sRig += sNum;
                //sRig += DateTime.Now.ToString("HHmmss");
                //sRig += "01";
                //sRig += "ZZZP";
                //sw.Write(sRig + _clsDef.CRLF);

                //sRig = "3";
                //sRig += sNeg.PadLeft(6, Convert.ToChar("0"));
                //sRig += "000000";
                //sRig += " ";
                //sRig += "  ";
                //sw.Write(sRig + _clsDef.CRLF);

                //sRig = "400000000       ";
                //sw.Write(sRig + _clsDef.CRLF);

                foreach (DataRow y in tabOrd.Rows)
                {
                    b = true;
                    //if (("" + y["orf_art"]).Trim() == "")
                    //    b = false;

                    if (((string)y["orf_ean"]).Trim() == "")
                    {
                        j = tEan.Select("ean_art='" + (string)y["orf_art"] + "'");
                        if(j.Length > 0)
                            y["orf_ean"] = (string)j[0]["ean_ean"];
                    }
                    if (((string)y["orf_ean"]).Trim() == "")
                        b = false;

                    if (b)
                    {
                        sRig = ((string)y["orf_ean"]).PadRight(13, Convert.ToChar(' '));

                        s = (Convert.ToInt16(y["orf_qta"])).ToString("0000.00");
                        sRig += s.Replace(",",".");

                        sw.Write(sRig + _clsDef.CRLF);
                    }

                }

                ((TextWriter)sw).Flush();
                sw.Close();
                sw.Dispose();

            }

            return sFil;
        }

        private void importazioneDaTerminalinoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ImpTerm();
        }

        private void ImpTerm()
        {
            string s = "";
            string sTer = "";
            DataTable tOrd = ((DataView)dgv1.DataSource).Table;
            if(tOrd.Rows.Count > 0)
            {
                if ((string)tOrd.Rows[tOrd.Rows.Count - 1]["orf_art"] == "")
                    tOrd.Rows[tOrd.Rows.Count - 1].Delete();
            }

            string sDivTip = "";
            s = "SELECT * FROM TabForImport WHERE tab_cod='" + _strOftFor + "'";
            DataTable t = _clsFun.FillTabSql(TABTABFIM, s, false, _strConSql);
            if (t.Rows.Count == 0)
                MessageBox.Show("Tabella importazione fornitore non configurata!");
            else if (DBNull.Value.Equals(t.Rows[0]["tab_tip"]) || ((string)t.Rows[0]["tab_tip"]).Trim() == "")
                MessageBox.Show("Tipo divulgazione non configurato in Tabella importazione fornitore!");
            else
            {
                sDivTip = (string)t.Rows[0]["tab_tip"];

                int i = tOrd.Rows.Count;

                frmGesImpTerm f = new frmGesImpTerm();
                f._strDivTip = sDivTip;
                f._strTip = "ORD";
                f.ShowDialog();
                if (f._tabImp != null)
                {
                    sTer = f._strTer;

                    foreach (DataRow y in f._tabImp.Rows)
                    {
                        if ((string)y["ord_art"] == "001261")
                            Console.WriteLine("aaaaa");
                        i++;
                        if (((string)y["ord_ean"]).Trim() != "")
                        {
                            s = Convert.ToInt64(y["ord_ean"]).ToString();
                            t = _clsQry.ArtSeek(s, "");
                            if (t != null && t.Rows.Count > 0)
                            {
                                DataRow x = tOrd.NewRow();
                                x["orf_nri"] = i.ToString("000");
                                x["orf_yea"] = _strOftYea;
                                x["orf_num"] = _strOftNum;
                                x["orf_day"] = DateTime.Today;
                                x["orf_ean"] = y["ord_ean"];
                                x["orf_qta"] = (decimal)y["ord_qta"];
                                x["orf_msg"] = (string)y["ord_msg"];
                                tOrd.Rows.Add(x);
                                FillOrfRow(t, x, false);
                            }
                            else
                            {
                                DataRow x = tOrd.NewRow();
                                x["orf_nri"] = i.ToString("000");
                                x["orf_yea"] = _strOftYea;
                                x["orf_num"] = _strOftNum;
                                x["orf_day"] = DateTime.Today;
                                x["orf_ean"] = y["ord_ean"];
                                x["orf_qta"] = (decimal)y["ord_qta"];
                                x["orf_msg"] = (string)y["ord_msg"] + " Barcode non trovato " + (string)y["ord_ean"];
                                tOrd.Rows.Add(x);
                            }
                        }
                        else if (((string)y["ord_art"]).Trim() != "")
                        {
                            if(sTer == _clsDef.TERMTDIV)
                                t = _clsQry.ArtSeek(((string)y["ord_art"]).Trim(), _strOftFor);
                            else
                                t = _clsQry.ArtSeek(((string)y["ord_art"]).Trim(), "ARF" + _strOftFor);

                            if (t != null && t.Rows.Count > 0)
                            {
                                DataRow x = tOrd.NewRow();
                                x["orf_nri"] = i.ToString("000");
                                x["orf_yea"] = _strOftYea;
                                x["orf_num"] = _strOftNum;
                                x["orf_day"] = DateTime.Today;
                                x["orf_arf"] = y["ord_art"];
                                x["orf_ean"] = "";
                                x["orf_qta"] = (decimal)y["ord_qta"];
                                tOrd.Rows.Add(x);
                                FillOrfRow(t, x, false);
                            }
                            else
                            {
                                DataRow x = tOrd.NewRow();
                                x["orf_nri"] = i.ToString("000");
                                x["orf_yea"] = _strOftYea;
                                x["orf_num"] = _strOftNum;
                                x["orf_day"] = DateTime.Today;
                                x["orf_arf"] = y["ord_art"];
                                x["orf_qta"] = (decimal)y["ord_qta"];
                                x["orf_msg"] = "Articolo fornitore non trovato " + (string)y["ord_art"];
                                tOrd.Rows.Add(x);
                            }
                        }

                        else
                        {
                            DataRow x = tOrd.NewRow();
                            x["orf_nri"] = i.ToString("000");
                            x["orf_yea"] = _strOftYea;
                            x["orf_num"] = _strOftNum;
                            x["orf_day"] = DateTime.Today;
                            x["orf_ean"] = y["ord_ean"];
                            x["orf_qta"] = (decimal)y["ord_qta"];
                            x["orf_msg"] = "Articolo non trovato" + " " + (string)y["ord_art"];
                            tOrd.Rows.Add(x);
                        }
                    }

                    Salva();

                    Totali();

                    if (tOrd.Rows.Count > 0)
                        NewRiga(false);
                }
            }
        }

        private void tuttoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OrdPdf("ALL");
        }

        private void erroriToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OrdPdf("ERR");
        }

        private void OrdPdf(string strTip)
        {
            string s = "";
            int iRows = 37;
            int iRow = 0;

            DataTable t = ((DataView)dgv1.DataSource).ToTable().Clone();
            if (strTip == "ALL")
            {
                DataTable tTmp = ((DataView)dgv1.DataSource).ToTable();

                foreach(DataRow y in tTmp.Rows)
                {
                    if ((decimal)y["orf_qta"] > 0 && ((string)y["orf_art"]).Trim() != "" && ((string)y["orf_msg"]).Trim() == "")
                        t.ImportRow(y);
                }

            }
            else if (strTip == "ERR")
            {
                DataView v = new DataView(((DataView)dgv1.DataSource).ToTable(), "orf_msg <>''", "", DataViewRowState.CurrentRows);
                t = v.ToTable();
            }

            // Create a new PDF document
            PdfDocument pd = new PdfDocument();
            pd.Info.Title = "Articoli";

            // Create an empty page
            PdfPage page = pd.AddPage();

            // Get an XGraphics object for drawing
            XGraphics gfx = XGraphics.FromPdfPage(page);

            //XPdfFontOptions options = new XPdfFontOptions(PdfFontEncoding.Unicode, PdfFontEmbedding.Always);

            // Create a font
            XFont font1 = new XFont("Courier new", 10, XFontStyle.Bold);
            XFont font2 = new XFont("Courier new", 9, XFontStyle.Regular);
            XFont font3 = new XFont("Courier new", 8, XFontStyle.Regular);
            XFont font4 = new XFont("Courier new", 8, XFontStyle.Bold);

            double Y = 5;
            double X = 20;

            XStringFormat frmDX = new XStringFormat();
            frmDX.Alignment = XStringAlignment.Far;
            frmDX.LineAlignment = XLineAlignment.Far;

            for (int i = 0; i <= t.Rows.Count - 1; i++)
            {

                if (iRow > iRows || iRow == 0)
                {
                    if (iRow > 0)
                    {
                        page = pd.AddPage();
                        gfx = XGraphics.FromPdfPage(page);
                    }

                    X = 20;
                    iRow = 0;
                    s = DateTime.Now.ToString("dd/MM/yyyy") + " STAMPA ORDINE " + lblOftNum.Text +"/" + lblOftYea.Text + " a " + cmbOftFor.Text;
                    //if (chkCli.Checked)
                    //    s += "CLIENTI ";
                    //if (chkFor.Checked)
                    //    s += "FORNITORI ";
                    gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    X += 20;

                    s = "Codice";
                    gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

                    s = "C.Forn";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 50, X, XStringFormats.Default);

                    s = "Barcode";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 100, X, XStringFormats.Default);

                    s = "Descrizione";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 180, X, XStringFormats.Default);

                    s = "Costo";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 440, X, XStringFormats.Default);

                    s = "Q.ta";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 490, X, XStringFormats.Default);

                    s = "Importo";
                    gfx.DrawString(s, font1, XBrushes.Black, Y + 530, X, XStringFormats.Default);

                    X += 15;
                }

                iRow++;

                if (t.Rows.Count - 1 >= i)
                {
                    s = "";
                    if (!DBNull.Value.Equals(t.Rows[i]["orf_art"]))
                        s = (string)t.Rows[i]["orf_art"];
                    gfx.DrawString(s, font3, XBrushes.Black, Y, X, XStringFormats.Default);

                    s = "";
                    if (!DBNull.Value.Equals(t.Rows[i]["orf_arf"]))
                        s = (string)t.Rows[i]["orf_arf"];
                    gfx.DrawString(s, font3, XBrushes.Black, Y + 50, X, XStringFormats.Default);

                    s = "";
                    if (!DBNull.Value.Equals(t.Rows[i]["orf_ean"]))
                        s = (string)t.Rows[i]["orf_ean"];
                    gfx.DrawString(s, font3, XBrushes.Black, Y + 100, X, XStringFormats.Default);

                    s = "";
                    if (!DBNull.Value.Equals(t.Rows[i]["orf_ard"]))
                        s = (string)t.Rows[i]["orf_ard"];
                    gfx.DrawString(s, font3, XBrushes.Black, Y + 180, X, XStringFormats.Default);

                    s = "";
                    if (!DBNull.Value.Equals(t.Rows[i]["orf_cos"]))
                        s = ((decimal)t.Rows[i]["orf_cos"]).ToString("####,##0.0000");
                    gfx.DrawString(s, font3, XBrushes.Black, Y + 470, X + 3, frmDX);

                    s = "";
                    if (!DBNull.Value.Equals(t.Rows[i]["orf_qta"]))
                        s = ((decimal)t.Rows[i]["orf_qta"]).ToString("####,##0.00");
                    gfx.DrawString(s, font3, XBrushes.Black, Y + 520, X + 3, frmDX);

                    s = "";
                    if (!DBNull.Value.Equals(t.Rows[i]["OrfImp"]))
                        s = ((decimal)t.Rows[i]["OrfImp"]).ToString("####,##0.00");
                    gfx.DrawString(s, font3, XBrushes.Black, Y + 570, X + 3, frmDX);

                    s = "";
                    if (!DBNull.Value.Equals(t.Rows[i]["orf_msg"]))
                        s = (string)t.Rows[i]["orf_msg"];
                    gfx.DrawString(s, font4, XBrushes.Black, Y+30, X+7, XStringFormats.Default);

                    //break;
                }

                X += 20;
                //break;

            }

            X += 15;

            //s = "Saldo    " + txtTotTot.Text;
            //gfx.DrawString(s, font1, XBrushes.Black, Y, X, XStringFormats.Default);

            // Save the document...
            try
            {
                string sFil = "C:\\ApProject\\PDF\\Ord_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                pd.Save(sFil);
                // ...and start a viewer.
                Process.Start(sFil);
            }
            catch (Exception ex)
            {
                MessageBox.Show("File di stampa già aperto!");
            }

        }

        private string InvioFtp(DataTable tabDiv, string strFil)
        {
            string sPathLoc = Path.GetDirectoryName(strFil)+"\\";
            string sFil = strFil;

            string s = ((string)tabDiv.Rows[0]["tab_hst"]).Trim();

            string[] a = s.Split('|');

            //string strHst = ((string)tabDiv.Rows[0]["tab_hst"]).Trim();

            string strHst = a[0];
            if( ((string)tabDiv.Rows[0]["tab_tip"]).Trim() == _clsDef.DIVMIGROSS)
                strHst = a[1];
            else if( ((string)tabDiv.Rows[0]["tab_tip"]).Trim() == _clsDef.DIVBRENDO)
                strHst = a[1];
            //{
            //    Console.WriteLine("xxxxx");
            //    s = ((string)tabDiv.Rows[0]["tab_hst"]).Trim();
            //    a = s.Split('|');
            //    Console.WriteLine("xxxxx");
            //    strHst = a[0];
            //}

            string strUsr = ((string)tabDiv.Rows[0]["tab_uso"]).Trim(); 
            string strPwd = ((string)tabDiv.Rows[0]["tab_pwo"]).Trim();
            string strPath = ((string)tabDiv.Rows[0]["tab_fpt"]).Trim();

            if( ((string)tabDiv.Rows[0]["tab_tip"]).Trim() == _clsDef.DIVMIGROSS)
            {
                a = strPath.Split('|');
                strPath = a[1];
            }

            string sMsg = "";
            clsFtp clsFtp = new clsFtp();
            clsFtp._strFtpHost = strHst;
            clsFtp._strFtpUser = strUsr;
            clsFtp._strFtpPswd = strPwd;
            clsFtp._strFtpPath = strPath;
            clsFtp._strLocPath = sPathLoc;
            clsFtp.FtpUpLoad(sFil);

            if (File.Exists(sPathLoc + "Temp\\" + Path.GetFileName(sFil)))
                File.Delete(sPathLoc + "Temp\\" + Path.GetFileName(sFil));

            clsFtp._strLocPath = sPathLoc + "Temp\\";
            sMsg += clsFtp.FtpDownLoad(sFil);
            FileInfo fi1 = new FileInfo(sPathLoc + "Temp\\" + Path.GetFileName(sFil));
            FileInfo fi2 = new FileInfo(sPathLoc + Path.GetFileName(sFil));
            if (sMsg == "")
            {
                if (fi1.Length != fi2.Length)
                    sMsg += sFil + " non inviato correttamente!";
            }
            s = Path.GetDirectoryName(sPathLoc) + "\\Save\\" + Path.GetFileName(sFil) + "_Inviato_" + DateTime.Now.ToString("yyyyMMddmmss");
            File.Move(sFil, s);

            return sMsg;
        }

        private void Invio2ApMatch(string strPth, string strFil)
        {
            string s = "";

            string sFil = Path.GetFileName(strFil);

            string sFilDest = strPth + sFil;
            string sFilTmp = strPth + sFil + ".tmp";

            if (!File.Exists(sFilDest))
                File.Copy(strFil, sFilDest);
            else
            {
                File.Move(sFilDest, sFilTmp);

                StreamWriter sw = new StreamWriter(sFilTmp, true);

                using (StreamReader sr = new StreamReader(strFil, System.Text.Encoding.Default))
                {
                    string sRig = "";

                    while ((sRig = sr.ReadLine()) != null)
                    {
                        sw.Write(sRig + _clsDef.CRLF);
                    }

                    sr.Close();
                    sr.Dispose();
                }

                sw.Close();
                sw.Dispose();

                File.Move(sFilTmp, sFilDest);

            }

            string sOld = Path.GetDirectoryName( strFil) +"\\old\\" + sFil + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");

            File.Move(strFil, sOld);

            Console.WriteLine("aaaaaa");

        }
    }
}
