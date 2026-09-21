using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Text;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows.Forms;

namespace APOffice
{

public partial class frmGesVariazioni : Form
{
	private class LookupCache
	{
		public Dictionary<string, DataRow> Reparti = new Dictionary<string, DataRow>();

		public Dictionary<string, DataRow> Iva = new Dictionary<string, DataRow>();

		public Dictionary<string, DataRow> OffArt = new Dictionary<string, DataRow>();

		public Dictionary<string, DataRow> OffTes = new Dictionary<string, DataRow>();

		public Dictionary<string, DataRow> Tgr = new Dictionary<string, DataRow>();

		public Dictionary<string, DataRow> Ori = new Dictionary<string, DataRow>();

		public Dictionary<string, DataRow> Cal = new Dictionary<string, DataRow>();

		public Dictionary<string, DataRow> Cat = new Dictionary<string, DataRow>();
	}

	private clsDefine _clsDef = new clsDefine();

	private clsFuncs _clsFun = new clsFuncs();

	private clsQuery _clsQry = new clsQuery();

	private clsVariazioni _clsVar = new clsVariazioni();

	private const string TABANAART = "AnaArticoli";

	private const string TABGESVAR = "GesVariazioni";

	private const string TABETIVAR = "GesVarEtichette";

	private const string TABANAOFT = "GesOffTestate";

	private const string TABANAOFF = "GesOffArticoli";

	private const string TABTABTGR = "TabTipoGrammatura";

	private const string TABLISACQ = "GesLisAcquisto";

	private const string TABLISVEN = "GesLisVendita";

	private const string TABTABSTD = "TabStatoDivulgazioni";

	private const string TABTABSTA = "TabStato";

	private const string TABTABOFF = "TabOfferte";

	private const string TABTABIVA = "TabIva";

	private const string TABTABREP = "TabReparti";

	private const string TABLISNEG = "TabLisNegozi";

	private const string INVIOTERM = "Invio a terminalino";

	private string _strConSql = "";

	private string _strPosCod = "";

	private string _strCnfPar = "";

	private string _strPar039SoloPrezziNeg2Pos = "";

	private bool _isProcessingEti = false;

	public string _strSta = "";

	private Dictionary<string, DataRow[]> _cacheArtEanRows = new Dictionary<string, DataRow[]>();

	private DataTable _cacheMasterTable = null;

	public frmGesVariazioni()
	{
		InitializeComponent();
		new clsGesGraph().SetGraph(this, 0);
		DoubleBuffered(dgv1, setting: true);
		DoubleBuffered(dgv2, setting: true);
		DoubleBuffered(dgv3, setting: true);
	}

	private new void DoubleBuffered(DataGridView dgv, bool setting)
	{
		Type dgvType = dgv.GetType();
		PropertyInfo pi = dgvType.GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic);
		if (pi != null)
		{
			pi.SetValue(dgv, setting, null);
		}
	}

	private void frmGesVariazioni_Load(object sender, EventArgs e)
	{
		Show();
		menuStrip1.Enabled = false;
		_strPar039SoloPrezziNeg2Pos = _clsFun.ParGet(clsDefine.enuParametri.Par039SoloPrezziNeg2Pos, _strConSql);
		_clsQry._strPar031Lotti2Pos = _clsFun.ParGet(clsDefine.enuParametri.Par031Lotti2Pos, _strConSql);
		if (_strSta == "NOVAR")
		{
			invioVariazioniACasseEBilanceToolStripMenuItem.Enabled = false;
			stampaEtichetteToolStripMenuItem.Enabled = false;
		}
		else
		{
			string s = _clsFun.FileIni("R", clsDefine.enuIni.Ini07PathDivCasse, "");
			if (s == "")
			{
				invioVariazioniACasseEBilanceToolStripMenuItem.Enabled = false;
			}
		}
		_strCnfPar = _clsQry.DefCnfPar("POS");
		_strConSql = _clsFun.ConSql("");
		SetDgv1();
		SetDgv2();
		SetDgv3();
		FillTab();
		Controlli();
		FillLisPromo();
		try
		{
			int countP = 0, countE = 0;
			_clsVar.ConsolidaOfferteAllAvvio(out countP, out countE);
		}
		catch (Exception ex)
		{
			_clsFun.ErrorLog("frmGesVariazioni.ConsolidaOfferte", ex.Message);
		}
		FillVarPos("", "");
		FillOfferte();
		InizializzaIconeMenu();
		menuStrip1.Enabled = true;
	}

	private void InizializzaIconeMenu()
	{
		try
		{
			esciToolStripMenuItem.Image = GetIcon("ðŸšª", Color.DarkRed);
			stampaEtichetteToolStripMenuItem.Image = GetIcon("ðŸ–\u00a8ï\u00b8\u008f", Color.SteelBlue);
			invioVariazioniACasseEBilanceToolStripMenuItem.Image = GetIcon("ðŸ“¤", Color.ForestGreen);
			utilityToolStripMenuItem.Image = GetIcon("âš™ï\u00b8\u008f", Color.DimGray);
			etichetteToolStripMenuItem.Image = GetIcon("ðŸ\u008f·ï\u00b8\u008f", Color.DarkOrange);
			annullaTutteLeEtichetteToolStripMenuItem.Image = GetIcon("â\u009dŒ", Color.Crimson);
			attivaTutteLeEtichetteToolStripMenuItem1.Image = GetIcon("âœ…", Color.MediumSeaGreen);
			recuperoEtichetteToolStripMenuItem1.Image = GetIcon("ðŸ”„", Color.RoyalBlue);
			casseEBilanceToolStripMenuItem.Image = GetIcon("ðŸ›’", Color.DarkSlateBlue);
			annullaTutteLeVariazioniToolStripMenuItem.Image = GetIcon("â\u009dŒ", Color.Crimson);
			attivaTutteLeVariazioniToolStripMenuItem.Image = GetIcon("âœ…", Color.MediumSeaGreen);
			recuperoVariazioniToolStripMenuItem.Image = GetIcon("ðŸ”„", Color.RoyalBlue);
			terminalinoToolStripMenuItem.Image = GetIcon("ðŸ“±", Color.Teal);
			impiantoCassaToolStripMenuItem.Image = GetIcon("ðŸ’°", Color.Sienna);
			bilanciaToolStripMenuItem.Image = GetIcon("âš–ï\u00b8\u008f", Color.Sienna);
			importTerminalinoToolStripMenuItem.Image = GetIcon("ðŸ“¥", Color.Goldenrod);
			importTerminalinoToolStripMenuItem1.Image = GetIcon("ðŸ“¥", Color.Goldenrod);
			refreshToolStripMenuItem.Image = GetIcon("ðŸ”„", Color.CornflowerBlue);
			promoToolStripMenuItem.Image = GetIcon("â\u00ad\u0090", Color.Gold);
			designerEtichetteToolStripMenuItem.Image = GetIcon("ðŸ–Œï\u00b8\u008f", Color.MediumPurple);
			svuotaElencoVariazioniToolStripMenuItem.Image = GetIcon("ðŸ—‘ï\u00b8\u008f", Color.Crimson);
		}
		catch
		{
		}
	}

	private Image GetIcon(string unicodeChar, Color fallbackColor)
	{
		Bitmap bmp = new Bitmap(24, 24);
		using (Graphics g = Graphics.FromImage(bmp))
		{
			g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
			using (Font f = new Font("Segoe UI Emoji", 12f, FontStyle.Regular))
			using (SolidBrush b = new SolidBrush(fallbackColor))
			{
				StringFormat fmt = new StringFormat();
				fmt.Alignment = StringAlignment.Center;
				fmt.LineAlignment = StringAlignment.Center;
				g.DrawString(unicodeChar, f, b, new RectangleF(0f, 0f, 24f, 24f), fmt);
			}
		}
		return bmp;
	}

	private void frmGesVariazioni_KeyDown(object sender, KeyEventArgs e)
	{
		if (e.KeyCode == Keys.Escape)
		{
			e.Handled = true;
			Esci();
		}
	}

	private void esciToolStripMenuItem_Click(object sender, EventArgs e)
	{
		Esci();
	}

	private void Esci()
	{
		try
		{
			Cursor oldCursor = Cursor;
			Cursor = Cursors.WaitCursor;
			Application.DoEvents();
			Salva();
		}
		catch (Exception ex)
		{
			_clsFun.ErrorLog("frmGesVariazioni.Esci", ex.Message);
		}
		finally
		{
			if (base.Owner != null)
			{
				try
				{
					base.Owner.BringToFront();
					base.Owner.Activate();
				}
				catch
				{
				}
			}
			Close();
		}
	}

	private void SetDgv1()
	{
		DataTable tSta = _clsFun.FillTabSql("TabStatoDivulgazioni", "SELECT * FROM TabStatoDivulgazioni WHERE tab_var=1", bPrimo: false, _strConSql);
		dgv1.AutoGenerateColumns = false;
		dgv1.AllowUserToAddRows = false;
		dgv1.ReadOnly = false;
		dgv1.AllowUserToDeleteRows = false;
		DataGridViewComboBoxColumn cCmb = new DataGridViewComboBoxColumn();
		cCmb.DataPropertyName = "pos_inv";
		cCmb.Name = "Stato";
		cCmb.Width = 60;
		cCmb.DataSource = tSta;
		cCmb.ValueMember = "tab_cod";
		cCmb.DisplayMember = "tab_des";
		cCmb.ReadOnly = true;
		cCmb.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing;
		cCmb.DefaultCellStyle.Font = new Font("Microsoft Sans", 8.75f, GraphicsUnit.Pixel);
		dgv1.Columns.Add(cCmb);
		DataGridViewTextBoxColumn cTbc = new DataGridViewTextBoxColumn();
		cTbc.DataPropertyName = "pos_tvd";
		cTbc.Name = "Tipo variazione";
		cTbc.Width = 50;
		cTbc.ValueType = typeof(string);
		cTbc.ReadOnly = true;
		dgv1.Columns.Add(cTbc);
		cTbc = new DataGridViewTextBoxColumn();
		cTbc.DataPropertyName = "pos_art";
		cTbc.Name = "Articolo";
		cTbc.Width = 60;
		cTbc.ValueType = typeof(string);
		cTbc.ReadOnly = true;
		dgv1.Columns.Add(cTbc);
		cTbc = new DataGridViewTextBoxColumn();
		cTbc.DataPropertyName = "pos_arb";
		cTbc.Name = "Descrizione cassa";
		cTbc.Width = 100;
		cTbc.ValueType = typeof(string);
		cTbc.ReadOnly = true;
		cTbc.ToolTipText = "Doppio click per anagrafica articolo";
		dgv1.Columns.Add(cTbc);
		cTbc = new DataGridViewTextBoxColumn();
		cTbc.DataPropertyName = "pos_day";
		cTbc.Name = "Data di modifica";
		cTbc.Width = 55;
		cTbc.ValueType = typeof(DateTime);
		cTbc.ReadOnly = true;
		dgv1.Columns.Add(cTbc);
		dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "dd/MM/yy";
		cTbc = new DataGridViewTextBoxColumn();
		cTbc.DataPropertyName = "pos_prv";
		cTbc.Name = "Prezzo vendita";
		cTbc.Width = 50;
		cTbc.ValueType = typeof(string);
		cTbc.ReadOnly = true;
		dgv1.Columns.Add(cTbc);
		dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
		dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "##0.00";
		cTbc = new DataGridViewTextBoxColumn();
		cTbc.DataPropertyName = "pos_pvo";
		cTbc.Name = "Prezzo vecchio";
		cTbc.Width = 50;
		cTbc.ValueType = typeof(string);
		cTbc.ReadOnly = true;
		dgv1.Columns.Add(cTbc);
		dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
		dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "##0.00";
		cTbc = new DataGridViewTextBoxColumn();
		cTbc.DataPropertyName = "pos_pxc";
		cTbc.Name = "Costo";
		cTbc.Width = 50;
		cTbc.ValueType = typeof(string);
		cTbc.ReadOnly = true;
		dgv1.Columns.Add(cTbc);
		dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
		dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "##0.000";
		cTbc = new DataGridViewTextBoxColumn();
		cTbc.DataPropertyName = "pos_mrg";
		cTbc.Name = "Margine %";
		cTbc.Width = 50;
		cTbc.ValueType = typeof(string);
		cTbc.ReadOnly = true;
		dgv1.Columns.Add(cTbc);
		dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
		dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "##0.00";
		cTbc = new DataGridViewTextBoxColumn();
		cTbc.DataPropertyName = "pos_ean";
		cTbc.Name = "Barcode";
		cTbc.Width = 90;
		cTbc.ValueType = typeof(string);
		cTbc.ReadOnly = true;
		dgv1.Columns.Add(cTbc);
		DataGridViewCheckBoxColumn cCbc = new DataGridViewCheckBoxColumn();
		cCbc.ValueType = typeof(bool);
		cCbc.DataPropertyName = "pos_eaa";
		cCbc.Name = "Ean Annullato";
		cCbc.Width = 30;
		cTbc.ReadOnly = true;
		dgv1.Columns.Add(cCbc);
		cTbc = new DataGridViewTextBoxColumn();
		cTbc.DataPropertyName = "pos_eaq";
		cTbc.Name = "Q.ta EAN";
		cTbc.Width = 30;
		cTbc.ValueType = typeof(string);
		cTbc.ReadOnly = true;
		dgv1.Columns.Add(cTbc);
		dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
		dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "##0";
		cTbc = new DataGridViewTextBoxColumn();
		cTbc.DataPropertyName = "pos_eap";
		cTbc.Name = "Prezzo EAN";
		cTbc.Width = 50;
		cTbc.ValueType = typeof(string);
		cTbc.ReadOnly = true;
		dgv1.Columns.Add(cTbc);
		dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
		dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "##0.00";
		cTbc = new DataGridViewTextBoxColumn();
		cTbc.DataPropertyName = "pos_pun";
		cTbc.Name = "Punti";
		cTbc.Width = 20;
		cTbc.ValueType = typeof(string);
		cTbc.ReadOnly = true;
		dgv1.Columns.Add(cTbc);
		dgv1.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
		dgv1.Columns[cTbc.Name].DefaultCellStyle.Format = "#0";
		cCbc = new DataGridViewCheckBoxColumn();
		cCbc.ValueType = typeof(bool);
		cCbc.DataPropertyName = "bil_bil";
		cCbc.Name = "Bilancia";
		cCbc.Width = 50;
		cTbc.ReadOnly = true;
		dgv1.Columns.Add(cCbc);
		cTbc = new DataGridViewTextBoxColumn();
		cTbc.DataPropertyName = "pos_ofd";
		cTbc.Name = "Offerta";
		cTbc.Width = 30;
		cTbc.ValueType = typeof(string);
		cTbc.ReadOnly = true;
		dgv1.Columns.Add(cTbc);
		cTbc = new DataGridViewTextBoxColumn();
		cTbc.DataPropertyName = "pos_msg";
		cTbc.Name = "Errori";
		cTbc.Width = 100;
		cTbc.ValueType = typeof(string);
		cTbc.ReadOnly = true;
		dgv1.Columns.Add(cTbc);
		cTbc = new DataGridViewTextBoxColumn();
		cTbc.DataPropertyName = "pos_ori";
		cTbc.Name = "Variazione";
		cTbc.Width = 100;
		cTbc.ValueType = typeof(string);
		cTbc.ReadOnly = true;
		dgv1.Columns.Add(cTbc);
		cTbc = new DataGridViewTextBoxColumn();
		cTbc.DataPropertyName = "pos_mix";
		cTbc.Name = "Mix-match";
		cTbc.Width = 30;
		cTbc.ValueType = typeof(string);
		cTbc.ReadOnly = true;
		dgv1.Columns.Add(cTbc);
	}

	private void SetDgv2()
	{
		dgv2.AutoGenerateColumns = false;
		dgv2.AllowUserToAddRows = false;
		dgv2.ReadOnly = false;
		dgv2.AllowUserToDeleteRows = false;
		DataGridViewTextBoxColumn cTbc = new DataGridViewTextBoxColumn();
		cTbc.DataPropertyName = "oft_yea";
		cTbc.Name = "Anno";
		cTbc.Width = 35;
		cTbc.ValueType = typeof(string);
		cTbc.ReadOnly = true;
		dgv2.Columns.Add(cTbc);
		cTbc = new DataGridViewTextBoxColumn();
		cTbc.DataPropertyName = "oft_cod";
		cTbc.Name = "Codice";
		cTbc.Width = 40;
		cTbc.ValueType = typeof(string);
		cTbc.ReadOnly = true;
		dgv2.Columns.Add(cTbc);
		cTbc = new DataGridViewTextBoxColumn();
		cTbc.DataPropertyName = "oft_des";
		cTbc.Name = "Descrizione";
		cTbc.Width = 140;
		cTbc.ValueType = typeof(string);
		cTbc.ReadOnly = true;
		dgv2.Columns.Add(cTbc);
		cTbc = new DataGridViewTextBoxColumn();
		cTbc.DataPropertyName = "oft_dti";
		cTbc.Name = "Data inizio";
		cTbc.Width = 70;
		cTbc.ValueType = typeof(string);
		cTbc.ReadOnly = true;
		dgv2.Columns.Add(cTbc);
		dgv2.Columns[cTbc.Name].DefaultCellStyle.Format = "dd/MM/yy";
		cTbc = new DataGridViewTextBoxColumn();
		cTbc.DataPropertyName = "oft_dtf";
		cTbc.Name = "Data fine";
		cTbc.Width = 70;
		cTbc.ValueType = typeof(string);
		cTbc.ReadOnly = true;
		dgv2.Columns.Add(cTbc);
		dgv2.Columns[cTbc.Name].DefaultCellStyle.Format = "dd/MM/yy";
		cTbc = new DataGridViewTextBoxColumn();
		cTbc.DataPropertyName = "OftStd";
		cTbc.Name = "Stato";
		cTbc.Width = 70;
		cTbc.ValueType = typeof(string);
		cTbc.ReadOnly = true;
		dgv2.Columns.Add(cTbc);
	}

	private void SetDgv3()
	{
		DataTable tSta = _clsFun.FillTabSql("TabStatoDivulgazioni", "SELECT * FROM TabStatoDivulgazioni WHERE tab_var=1", bPrimo: false, _strConSql);
		DataRow[] j = tSta.Select("tab_cod='0'");
		if (j.Length != 0)
		{
			j[0]["tab_des"] = "Da stampare";
		}
		j = tSta.Select("tab_cod='1'");
		if (j.Length != 0)
		{
			j[0]["tab_des"] = "Stampata";
		}
		dgv3.AutoGenerateColumns = false;
		dgv3.AllowUserToAddRows = false;
		dgv3.ReadOnly = false;
		dgv3.AllowUserToDeleteRows = false;
		DataGridViewComboBoxColumn cCmb = new DataGridViewComboBoxColumn();
		cCmb.DataPropertyName = "eti_inv";
		cCmb.Name = "Stato";
		cCmb.Width = 50;
		cCmb.DataSource = tSta.Copy();
		cCmb.ValueMember = "tab_cod";
		cCmb.DisplayMember = "tab_des";
		cCmb.ReadOnly = true;
		cCmb.DisplayStyle = DataGridViewComboBoxDisplayStyle.Nothing;
		cCmb.DefaultCellStyle.Font = new Font("Microsoft Sans", 8.75f, GraphicsUnit.Pixel);
		dgv3.Columns.Add(cCmb);
		DataGridViewTextBoxColumn cTbc = new DataGridViewTextBoxColumn();
		cTbc.DataPropertyName = "eti_art";
		cTbc.Name = "Articolo";
		cTbc.Width = 55;
		cTbc.ValueType = typeof(string);
		cTbc.ReadOnly = true;
		dgv3.Columns.Add(cTbc);
		cTbc = new DataGridViewTextBoxColumn();
		cTbc.DataPropertyName = "eti_ard";
		cTbc.Name = "Descrizione cassa";
		cTbc.Width = 150;
		cTbc.ValueType = typeof(string);
		cTbc.ReadOnly = true;
		cTbc.ToolTipText = "Doppio click per anagrafica articolo";
		dgv3.Columns.Add(cTbc);
		cTbc = new DataGridViewTextBoxColumn();
		cTbc.DataPropertyName = "eti_day";
		cTbc.Name = "Data di modifica";
		cTbc.Width = 55;
		cTbc.ValueType = typeof(DateTime);
		cTbc.ReadOnly = true;
		dgv3.Columns.Add(cTbc);
		dgv3.Columns[cTbc.Name].DefaultCellStyle.Format = "dd/MM/yy";
		cTbc = new DataGridViewTextBoxColumn();
		cTbc.DataPropertyName = "eti_qta";
		cTbc.Name = "Q.tà";
		cTbc.Width = 25;
		cTbc.ValueType = typeof(string);
		cTbc.ReadOnly = false;
		cTbc.MaxInputLength = 2;
		dgv3.Columns.Add(cTbc);
		dgv3.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
		dgv3.Columns[cTbc.Name].DefaultCellStyle.Format = "#0";
		cTbc = new DataGridViewTextBoxColumn();
		cTbc.DataPropertyName = "eti_prv";
		cTbc.Name = "Prezzo vendita";
		cTbc.Width = 45;
		cTbc.ValueType = typeof(string);
		cTbc.ReadOnly = true;
		dgv3.Columns.Add(cTbc);
		dgv3.Columns[cTbc.Name].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
		dgv3.Columns[cTbc.Name].DefaultCellStyle.Format = "##0.00";
		cTbc = new DataGridViewTextBoxColumn();
		cTbc.DataPropertyName = "eti_msg";
		cTbc.Name = "Errori";
		cTbc.Width = 100;
		cTbc.ValueType = typeof(string);
		cTbc.ReadOnly = true;
		dgv3.Columns.Add(cTbc);
		cTbc = new DataGridViewTextBoxColumn();
		cTbc.DataPropertyName = "eti_ori";
		cTbc.Name = "Variazione";
		cTbc.Width = 200;
		cTbc.ValueType = typeof(string);
		cTbc.ReadOnly = true;
		dgv3.Columns.Add(cTbc);
	}

	private void FillTab()
	{
		string p = "TabEtichette";
		string s = "SELECT * FROM " + p + " WHERE tab_ann=0 ORDER BY tab_cod";
		DataTable t = _clsFun.FillTabSql(p, s, bPrimo: false, _strConSql);
		ToolStripMenuItem item = (ToolStripMenuItem)menuStrip1.Items["stampaEtichetteToolStripMenuItem"];
		item.DropDownItems.Clear();
		foreach (DataRow y in t.Rows)
		{
			ToolStripMenuItem newItem = new ToolStripMenuItem((string)y["tab_cod"] + " " + y["tab_des"]);
			newItem.Click += eti_Click;
			item.DropDownItems.Add(newItem);
		}
		string s2 = "SELECT eti_cod as tab_cod, eti_des as tab_des FROM TabEtiFormatiLayout WHERE eti_ann=0 AND eti_tip='PDF' ORDER BY eti_cod";
		DataTable t2 = _clsFun.FillTabSql("TabEtiFormatiLayout", s2, bPrimo: false, _strConSql);
		foreach (DataRow y2 in t2.Rows)
		{
			ToolStripMenuItem newItem2 = new ToolStripMenuItem((string)y2["tab_cod"] + " " + y2["tab_des"]);
			newItem2.Click += eti_Click;
			item.DropDownItems.Add(newItem2);
		}
	}

	private void eti_Click(object sender, EventArgs e)
	{
		if (_isProcessingEti)
		{
			return;
		}
		_isProcessingEti = true;
		try
		{
			string s = "eti_inv='" + _clsDef.DIVDAD + "'";
			if (cmbEtiRep.SelectedValue != null && cmbEtiRep.SelectedValue.ToString() != "")
			{
				s = s + " AND eti_rep='" + cmbEtiRep.SelectedValue.ToString() + "'";
			}
			DataView v = new DataView((DataTable)dgv3.DataSource, s, "", DataViewRowState.CurrentRows);
			DataTable t = v.ToTable();
			clsGenPdfEtiDB.SanitizeUnitPrices(t);
			ToolStripMenuItem clickedItem = (ToolStripMenuItem)sender;
			string sCodItem = clickedItem.Text.Substring(0, 3);
			int codInt = 0;
			if (int.TryParse(sCodItem, out codInt) && codInt >= 900)
			{
				int skipLabels = Math.Max(0, (int)nudEtiRow.Value - 1);
				new clsGenPdfEtiDB().PrnPdfEtiDaDB(t, sCodItem, skipLabels);
				GestisciStatoPostStampa(v, sCodItem, isDbFormat: true);
				return;
			}
			if (sCodItem == "001")
			{
				string sCod = clickedItem.Text.Substring(0, 3);
				string sPar = "";
				s = "SELECT tab_par FROM TabEtichette WHERE tab_cod='" + sCod + "'";
				DataTable tEti = _clsFun.FillTabSql("TabEti", s, bPrimo: true, _strConSql);
				if (tEti.Rows.Count > 0)
				{
					sPar = (string)tEti.Rows[0]["tab_par"];
				}
				clsGenPdfEti obj = new clsGenPdfEti();
				obj.StartRow = (int)nudEtiRow.Value;
				obj.PrnPdfEti001(t, sPar);
			}
			if (clickedItem.Text.Substring(0, 3) == "002")
			{
				clsGenPdfEti obj2 = new clsGenPdfEti();
				obj2.StartRow = (int)nudEtiRow.Value;
				obj2.PrnPdfEti002(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "003")
			{
				clsGenPdfEti obj3 = new clsGenPdfEti();
				obj3.StartRow = (int)nudEtiRow.Value;
				obj3.PrnPdfEti003(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "004")
			{
				clsGenPdfEti obj4 = new clsGenPdfEti();
				obj4.StartRow = (int)nudEtiRow.Value;
				obj4.PrnPdfEti004(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "005")
			{
				clsGenPdfEti obj5 = new clsGenPdfEti();
				obj5.StartRow = (int)nudEtiRow.Value;
				obj5.PrnPdfEti005(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "006")
			{
				clsGenPdfEti obj6 = new clsGenPdfEti();
				obj6.StartRow = (int)nudEtiRow.Value;
				obj6.PrnPdfEti006(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "007")
			{
				clsGenPdfEti obj7 = new clsGenPdfEti();
				obj7.StartRow = (int)nudEtiRow.Value;
				obj7.PrnPdfEti007(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "008")
			{
				clsGenPdfEti obj8 = new clsGenPdfEti();
				obj8.StartRow = (int)nudEtiRow.Value;
				obj8.PrnPdfEti008(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "009")
			{
				clsGenPdfEti obj9 = new clsGenPdfEti();
				obj9.StartRow = (int)nudEtiRow.Value;
				obj9.PrnPdfEti009(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "010")
			{
				clsGenPdfEti obj10 = new clsGenPdfEti();
				obj10.StartRow = (int)nudEtiRow.Value;
				obj10.PrnPdfEti010(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "011")
			{
				clsGenPdfEti obj11 = new clsGenPdfEti();
				obj11.StartRow = (int)nudEtiRow.Value;
				obj11.PrnPdfEti011(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "012")
			{
				clsGenPdfEti obj12 = new clsGenPdfEti();
				obj12.StartRow = (int)nudEtiRow.Value;
				obj12.PrnPdfEti012(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "013")
			{
				clsGenPdfEti obj13 = new clsGenPdfEti();
				obj13.StartRow = (int)nudEtiRow.Value;
				obj13.PrnPdfEti013(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "014")
			{
				clsGenPdfEti obj14 = new clsGenPdfEti();
				obj14.StartRow = (int)nudEtiRow.Value;
				obj14.PrnPdfEti014(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "015")
			{
				clsGenPdfEti obj15 = new clsGenPdfEti();
				obj15.StartRow = (int)nudEtiRow.Value;
				obj15.PrnPdfEti015(t, nudEtiRow.Value.ToString());
			}
			if (clickedItem.Text.Substring(0, 3) == "016")
			{
				clsGenPdfEti obj16 = new clsGenPdfEti();
				obj16.StartRow = (int)nudEtiRow.Value;
				obj16.PrnPdfEti016(t, nudEtiRow.Value.ToString());
			}
			if (clickedItem.Text.Substring(0, 3) == "017")
			{
				clsGenPdfEti obj17 = new clsGenPdfEti();
				obj17.StartRow = (int)nudEtiRow.Value;
				obj17.PrnPdfEti017(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "020")
			{
				clsGenPdfEti obj18 = new clsGenPdfEti();
				obj18.StartRow = (int)nudEtiRow.Value;
				obj18.PrnPdfEti020(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "021")
			{
				clsGenPdfEti obj19 = new clsGenPdfEti();
				obj19.StartRow = (int)nudEtiRow.Value;
				obj19.PrnPdfEti021(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "022")
			{
				clsGenPdfEti obj20 = new clsGenPdfEti();
				obj20.StartRow = (int)nudEtiRow.Value;
				obj20.PrnPdfEti022(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "023")
			{
				clsGenPdfEti obj21 = new clsGenPdfEti();
				obj21.StartRow = (int)nudEtiRow.Value;
				obj21.PrnPdfEti023(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "024")
			{
				clsGenPdfEti obj22 = new clsGenPdfEti();
				obj22.StartRow = (int)nudEtiRow.Value;
				obj22.PrnPdfEti024(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "025")
			{
				clsGenPdfEti obj23 = new clsGenPdfEti();
				obj23.StartRow = (int)nudEtiRow.Value;
				obj23.PrnPdfEti025(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "026")
			{
				clsGenPdfEti obj24 = new clsGenPdfEti();
				obj24.StartRow = (int)nudEtiRow.Value;
				obj24.PrnPdfEti026(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "027")
			{
				clsGenPdfEti obj25 = new clsGenPdfEti();
				obj25.StartRow = (int)nudEtiRow.Value;
				obj25.PrnPdfEti027(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "028")
			{
				clsGenPdfEti obj26 = new clsGenPdfEti();
				obj26.StartRow = (int)nudEtiRow.Value;
				obj26.PrnPdfEti028(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "029")
			{
				clsGenPdfEti obj27 = new clsGenPdfEti();
				obj27.StartRow = (int)nudEtiRow.Value;
				obj27.PrnPdfEti029(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "030")
			{
				string sCod2 = clickedItem.Text.Substring(0, 3);
				string sPar2 = "";
				s = "SELECT tab_par FROM TabEtichette WHERE tab_cod='" + sCod2 + "'";
				DataTable tEti2 = _clsFun.FillTabSql("TabEti", s, bPrimo: true, _strConSql);
				if (tEti2.Rows.Count > 0)
				{
					sPar2 = (string)tEti2.Rows[0]["tab_par"];
				}
				clsGenPdfEti2 obj28 = new clsGenPdfEti2();
				obj28.StartRow = (int)nudEtiRow.Value;
				obj28.PrnPdfEti030(t, sPar2);
			}
			if (clickedItem.Text.Substring(0, 3) == "031")
			{
				clsGenPdfEti2 obj29 = new clsGenPdfEti2();
				obj29.StartRow = (int)nudEtiRow.Value;
				obj29.PrnPdfEti031(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "032")
			{
				clsGenPdfEti2 obj30 = new clsGenPdfEti2();
				obj30.StartRow = (int)nudEtiRow.Value;
				obj30.PrnPdfEti032(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "033")
			{
				clsGenPdfEti2 obj31 = new clsGenPdfEti2();
				obj31.StartRow = (int)nudEtiRow.Value;
				obj31.PrnPdfEti033(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "034")
			{
				clsGenPdfEti2 obj32 = new clsGenPdfEti2();
				obj32.StartRow = (int)nudEtiRow.Value;
				obj32.PrnPdfEti034(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "035")
			{
				clsGenPdfEti2 obj33 = new clsGenPdfEti2();
				obj33.StartRow = (int)nudEtiRow.Value;
				obj33.PrnPdfEti035(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "036")
			{
				clsGenPdfEti2 obj34 = new clsGenPdfEti2();
				obj34.StartRow = (int)nudEtiRow.Value;
				obj34.PrnPdfEti036(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "037")
			{
				clsGenPdfEti2 obj35 = new clsGenPdfEti2();
				obj35.StartRow = (int)nudEtiRow.Value;
				obj35.PrnPdfEti037(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "038")
			{
				clsGenPdfEti2 obj36 = new clsGenPdfEti2();
				obj36.StartRow = (int)nudEtiRow.Value;
				obj36.PrnPdfEti038(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "039")
			{
				clsGenPdfEti2 obj37 = new clsGenPdfEti2();
				obj37.StartRow = (int)nudEtiRow.Value;
				obj37.PrnPdfEti039(t, "");
			}
			if (clickedItem.Text.Substring(0, 3) == "040")
			{
				clsGenPdfEti2 obj38 = new clsGenPdfEti2();
				obj38.StartRow = (int)nudEtiRow.Value;
				obj38.PrnPdfEti040(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "041")
			{
				clsGenPdfEti2 obj39 = new clsGenPdfEti2();
				obj39.StartRow = (int)nudEtiRow.Value;
				obj39.PrnPdfEti041(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "042")
			{
				clsGenPdfEti2 obj40 = new clsGenPdfEti2();
				obj40.StartRow = (int)nudEtiRow.Value;
				obj40.PrnPdfEti042(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "043")
			{
				ImpEtiElettroniche(clickedItem.Text.Substring(0, 3));
			}
			if (clickedItem.Text.Substring(0, 3) == "044")
			{
				clsGenPdfEti2 obj41 = new clsGenPdfEti2();
				obj41.StartRow = (int)nudEtiRow.Value;
				obj41.PrnPdfEti044(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "045")
			{
				clsGenPdfEti2 obj42 = new clsGenPdfEti2();
				obj42.StartRow = (int)nudEtiRow.Value;
				obj42.PrnPdfEti045(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "046")
			{
				clsGenPdfEti2 obj43 = new clsGenPdfEti2();
				obj43.StartRow = (int)nudEtiRow.Value;
				obj43.PrnPdfEti046(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "047")
			{
				clsGenPdfEti2 obj44 = new clsGenPdfEti2();
				obj44.StartRow = (int)nudEtiRow.Value;
				obj44.PrnPdfEti047(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "048")
			{
				clsGenPdfEti2 obj45 = new clsGenPdfEti2();
				obj45.StartRow = (int)nudEtiRow.Value;
				obj45.PrnPdfEti048(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "049")
			{
				clsGenPdfEti2 obj46 = new clsGenPdfEti2();
				obj46.StartRow = (int)nudEtiRow.Value;
				obj46.PrnPdfEti049(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "050")
			{
				clsGenPdfEti2 obj47 = new clsGenPdfEti2();
				obj47.StartRow = (int)nudEtiRow.Value;
				obj47.PrnPdfEti050(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "051")
			{
				clsGenPdfEti2 obj48 = new clsGenPdfEti2();
				obj48.StartRow = (int)nudEtiRow.Value;
				obj48.PrnPdfEti051(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "052")
			{
				clsGenPdfEti2 obj49 = new clsGenPdfEti2();
				obj49.StartRow = (int)nudEtiRow.Value;
				obj49.PrnPdfEti052(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "053")
			{
				clsGenPdfEti2 obj50 = new clsGenPdfEti2();
				obj50.StartRow = (int)nudEtiRow.Value;
				obj50.PrnPdfEti053(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "054")
			{
				clsGenPdfEti2 obj51 = new clsGenPdfEti2();
				obj51.StartRow = (int)nudEtiRow.Value;
				obj51.PrnPdfEti054(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "055")
			{
				s = "1 Sotto Costo -";
				s += "2 1+1 - ";
				s = s + "3 P.Shock - " + _clsDef.CRLF;
				s += "4 Piac. Risparmio - ";
				s += "5 Super Offerta";
				frmUtyInput1 f = new frmUtyInput1();
				f._strDes = "Selezione immagine di stampa " + _clsDef.CRLF + s;
				f.ShowDialog(this);
				s = f._strTxt;
				if (_clsFun.Numerico(s, "12345"))
				{
					clsGenPdfEti2 obj52 = new clsGenPdfEti2();
					obj52.StartRow = (int)nudEtiRow.Value;
					obj52.PrnPdfEti055(t, s);
				}
			}
			if (clickedItem.Text.Substring(0, 3) == "056")
			{
				clsGenPdfEti2 obj53 = new clsGenPdfEti2();
				obj53.StartRow = (int)nudEtiRow.Value;
				obj53.PrnPdfEti056(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "057")
			{
				clsGenPdfEti2 obj54 = new clsGenPdfEti2();
				obj54.StartRow = (int)nudEtiRow.Value;
				obj54.PrnPdfEti057(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "058")
			{
				frmUtyInput1 f2 = new frmUtyInput1();
				f2._strDes = "Imputare il titolo dell'etichetta";
				f2.ShowDialog(this);
				s = f2._strTxt;
				if (s != "")
				{
					clsGenPdfEti2 obj55 = new clsGenPdfEti2();
					obj55.StartRow = (int)nudEtiRow.Value;
					obj55.PrnPdfEti058(t, s);
				}
			}
			if (clickedItem.Text.Substring(0, 3) == "059")
			{
				clsGenPdfEti2 obj56 = new clsGenPdfEti2();
				obj56.StartRow = (int)nudEtiRow.Value;
				obj56.PrnPdfEti059(t);
			}
			if (clickedItem.Text.Substring(0, 3) == "060")
			{
				s = "1 Prezzi bassi -";
				s += "2 Novità ";
				frmUtyInput1 f3 = new frmUtyInput1();
				f3._strDes = "Selezione immagine di stampa " + _clsDef.CRLF + s;
				f3.ShowDialog(this);
				s = f3._strTxt;
				if (_clsFun.Numerico(s, "12345"))
				{
					clsGenPdfEti2 obj57 = new clsGenPdfEti2();
					obj57.StartRow = (int)nudEtiRow.Value;
					obj57.PrnPdfEti060(t, s);
				}
			}
			if (clickedItem.Text.Substring(0, 3) == "061")
			{
				frmUtyInput1 f4 = new frmUtyInput1();
				f4._strDes = "Imputare il titolo dell'etichetta";
				f4.ShowDialog(this);
				s = f4._strTxt;
				if (s != "")
				{
					clsGenPdfEti2 obj58 = new clsGenPdfEti2();
					obj58.StartRow = (int)nudEtiRow.Value;
					obj58.PrnPdfEti061(t, s);
				}
			}
			if (clickedItem.Text.Substring(0, 3) == "018" || clickedItem.Text.Substring(0, 3) == "019")
			{
				string sCod3 = clickedItem.Text.Substring(0, 3);
				string sMsg = "";
				s = "SELECT tab_msg FROM TabEtichette WHERE tab_cod='" + sCod3 + "'";
				DataTable tMsg = _clsFun.FillTabSql("TabEti", s, bPrimo: true, _strConSql);
				if (tMsg.Rows.Count > 0)
				{
					s = (string)tMsg.Rows[0]["tab_msg"];
					string[] a = s.Split(';');
					sMsg = a[0];
					if (a.Length > 1)
					{
						frmUtyMsg1 f5 = new frmUtyMsg1();
						f5._strTip = "001";
						f5._strStr = s;
						f5.ShowDialog(this);
						sMsg = f5._strMsg;
					}
				}
				if (sCod3 == "018")
				{
					clsGenPdfEti obj59 = new clsGenPdfEti();
					obj59.StartRow = (int)nudEtiRow.Value;
					obj59.PrnPdfEti018(t, sMsg);
				}
				else if (sCod3 == "019")
				{
					clsGenPdfEti obj60 = new clsGenPdfEti();
					obj60.StartRow = (int)nudEtiRow.Value;
					obj60.PrnPdfEti019(t, sMsg);
				}
			}
			if (clickedItem.Text.Substring(0, 3) != "XXX")
			{
				GestisciStatoPostStampa(v, clickedItem.Text.Substring(0, 3), isDbFormat: false);
			}
		}
		finally
		{
			_isProcessingEti = false;
		}
	}

	private void GestisciStatoPostStampa(DataView v, string sCodItem, bool isDbFormat)
	{
		string sFlagPostStampa = "S";
		try
		{
			if (isDbFormat)
			{
				DataTable tFmt = _clsFun.FillTabSql("TabEtiFormatiLayout", "SELECT eti_sta FROM TabEtiFormatiLayout WHERE eti_cod='" + sCodItem + "'", bPrimo: true, _strConSql);
				if (tFmt != null && tFmt.Rows.Count > 0 && tFmt.Columns.Contains("eti_sta") && !DBNull.Value.Equals(tFmt.Rows[0]["eti_sta"]))
				{
					sFlagPostStampa = tFmt.Rows[0]["eti_sta"].ToString().Trim().ToUpper();
				}
			}
			else
			{
				DataTable tEtiTab = _clsFun.FillTabSql("TabEtichette", "SELECT tab_sta FROM TabEtichette WHERE tab_cod='" + sCodItem + "'", bPrimo: true, _strConSql);
				if (tEtiTab != null && tEtiTab.Rows.Count > 0 && tEtiTab.Columns.Contains("tab_sta") && !DBNull.Value.Equals(tEtiTab.Rows[0]["tab_sta"]))
				{
					sFlagPostStampa = tEtiTab.Rows[0]["tab_sta"].ToString().Trim().ToUpper();
				}
			}
		}
		catch { }

		if (string.IsNullOrEmpty(sFlagPostStampa))
		{
			sFlagPostStampa = "S";
		}

		if (sFlagPostStampa == "S")
		{
			if (MessageBox.Show(this, "Segnalare come stampate le etichette?", "STAMPA", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
			{
				foreach (DataRowView y in v)
				{
					y["eti_inv"] = _clsDef.DIVDIV;
				}
			}
		}
		else if (sFlagPostStampa == "E")
		{
			for (int i = v.Count - 1; i >= 0; i--)
			{
				v[i].Delete();
			}
		}
		// Se "N": Da Stampare -> non tocca lo stato e rimangono in lista (eti_inv = DIVDAD)
	}

	private void designerEtichetteToolStripMenuItem_Click(object sender, EventArgs e)
	{
		frmUtyEtiFormati f = new frmUtyEtiFormati();
		f._strConSql = _strConSql;
		f.ShowDialog(this);
		FillTab();
	}

	private void Controlli()
	{
		DataTable tCnf = _clsQry.ConfSeek("", "POS");
		_strPosCod = (string)tCnf.Rows[0]["cnf_pos"];
		string sMsg = "";
		if (_strPosCod == Convert.ToInt16(clsDefine.enuPos.posNcr745x).ToString("00"))
		{
			DataTable tIva = new clsQuery().TabIva();
			foreach (DataRow y in tIva.Rows)
			{
				if (!(bool)y["tab_ann"] && (DBNull.Value.Equals(y["tab_pos"]) || Convert.ToInt16(y["tab_pos"]) == 0))
				{
					sMsg = sMsg + "IVA " + ((string)y["tab_des"]).Trim() + " non impostata sulla voce NCR in tabella";
				}
			}
		}
		if (sMsg != "")
		{
			MessageBox.Show(this, sMsg, "CONTROLLI");
		}
	}

	private void CacheArticleRows(string sArt, DataRow[] rows)
	{
		if (!string.IsNullOrEmpty(sArt))
		{
			DataRow[] r = rows ?? new DataRow[0];
			string trimmed = sArt.Trim();
			string unpadded = trimmed.TrimStart('0');
			if (unpadded == "")
			{
				unpadded = "0";
			}
			string padded = unpadded.PadLeft(7, '0');
			if (r.Length != 0 || !TryGetCachedArticleRows(padded, out var existing) || existing == null || existing.Length == 0)
			{
				_cacheArtEanRows[sArt] = r;
				_cacheArtEanRows[trimmed] = r;
				_cacheArtEanRows[unpadded] = r;
				_cacheArtEanRows[padded] = r;
			}
		}
	}

	private bool TryGetCachedArticleRows(string sArt, out DataRow[] rows)
	{
		rows = null;
		if (string.IsNullOrEmpty(sArt))
		{
			return false;
		}
		if (_cacheArtEanRows.TryGetValue(sArt, out rows))
		{
			return true;
		}
		string trimmed = sArt.Trim();
		if (_cacheArtEanRows.TryGetValue(trimmed, out rows))
		{
			return true;
		}
		string unpadded = trimmed.TrimStart('0');
		if (unpadded == "")
		{
			unpadded = "0";
		}
		if (_cacheArtEanRows.TryGetValue(unpadded, out rows))
		{
			return true;
		}
		string padded = unpadded.PadLeft(7, '0');
		if (_cacheArtEanRows.TryGetValue(padded, out rows))
		{
			return true;
		}
		return false;
	}

	private void ClearArticleCache()
	{
		_cacheArtEanRows.Clear();
		if (_cacheMasterTable != null)
		{
			_cacheMasterTable.Dispose();
			_cacheMasterTable = null;
		}
	}

	private static void ClearColumnMaxLength(DataTable dt)
	{
		if (dt != null)
		{
			foreach (DataColumn col in dt.Columns)
			{
				if (col.DataType == typeof(string))
				{
					col.MaxLength = -1;
				}
			}
		}
	}

	private void FillVarPos(string strTip, string strVarNum)
	{
		ClearArticleCache();
		DataTable tPos = new clsGenTabTmp().TabTmpPos("PosArt");
		DataTable tEti = new clsGenTabTmp().TabTmpEti("EtiArt");
		DataTable tOff = _clsQry.OfsSeek(DateTime.Today);
		DataTable tOffEti = _clsQry.OfsSeekEti(DateTime.Today.AddDays(-15.0));
		string s = "SELECT * FROM TabOfferte WHERE tab_ann=0";
		DataTable tTof = _clsFun.FillTabSql("TabOfferte", s, bPrimo: false, _strConSql);
		s = "SELECT * FROM TabReparti WHERE tab_ann=0";
		DataTable tRep = _clsFun.FillTabSql("TabReparti", s, bPrimo: false, _strConSql);
		s = "SELECT * FROM TabIva WHERE tab_ann=0";
		DataTable tIva = _clsFun.FillTabSql("TabIva", s, bPrimo: false, _strConSql);
		DataTable tPro = _clsQry.LivPromo();
		DataTable tTgr = _clsFun.FillTabSql("TabTipoGrammatura", "SELECT * FROM TabTipoGrammatura", bPrimo: false, _strConSql);
		DataTable tCal = _clsFun.FillTabSql("TabCalibro", "SELECT * FROM TabCalibro", bPrimo: false, _strConSql);
		DataTable tCat = _clsFun.FillTabSql("TabCategoria", "SELECT * FROM TabCategoria", bPrimo: false, _strConSql);
		DataTable tOri = _clsFun.FillTabSql("TabOrigine", "SELECT * FROM TabOrigine", bPrimo: false, _strConSql);
		DataTable tLne = _clsFun.FillTabSql("TabLisNegozi", "SELECT tab_lis FROM TabNegozi WHERE tab_tip = 'L'", bPrimo: false, _strConSql);
		if (strTip != "")
		{
			if (dgv1.DataSource is DataTable dt1)
			{
				tPos = dt1;
			}
			if (dgv3.DataSource is DataTable dt3)
			{
				tEti = dt3;
			}
		}
		if (tPos == null)
		{
			tPos = new clsGenTabTmp().TabTmpPos("PosArt");
		}
		if (tEti == null)
		{
			tEti = new clsGenTabTmp().TabTmpEti("EtiArt");
		}
		DataTable templateEtiCheck = new clsGenTabTmp().TabTmpEti("EtiArt");
		foreach (DataColumn col in templateEtiCheck.Columns)
		{
			if (!tEti.Columns.Contains(col.ColumnName))
			{
				tEti.Columns.Add(new DataColumn(col.ColumnName, col.DataType)
				{
					Caption = col.Caption,
					DefaultValue = col.DefaultValue
				});
			}
		}
		DataTable templatePosCheck = new clsGenTabTmp().TabTmpPos("PosArt");
		foreach (DataColumn col2 in templatePosCheck.Columns)
		{
			if (!tPos.Columns.Contains(col2.ColumnName))
			{
				tPos.Columns.Add(new DataColumn(col2.ColumnName, col2.DataType)
				{
					Caption = col2.Caption,
					DefaultValue = col2.DefaultValue
				});
			}
		}
		ClearColumnMaxLength(tPos);
		ClearColumnMaxLength(tEti);
		DataSet dSet = new DataSet();
		Action<DataTable> safeAddTable = delegate(DataTable tbl)
		{
			if (tbl != null && !dSet.Tables.Contains(tbl.TableName))
			{
				if (tbl.DataSet != null)
				{
					tbl.DataSet.Tables.Remove(tbl);
				}
				dSet.Tables.Add(tbl);
			}
		};
		safeAddTable(tPos);
		safeAddTable(tOff);
		safeAddTable(tTof);
		safeAddTable(tRep);
		safeAddTable(tIva);
		safeAddTable(tPro);
		safeAddTable(tTgr);
		safeAddTable(tCal);
		safeAddTable(tCat);
		safeAddTable(tOri);
		safeAddTable(tLne);
		s = "SELECT var_tip, var_num, var_art, var_ori, var_dti, var_dtv, var_inv, var_etq, var_off, var_ean FROM GesVariazioni WHERE ";
		s = ((!(strVarNum == "")) ? (s + "var_tip = '" + strTip + "' AND var_num='" + strVarNum + "'") : (s + "var_num='" + _clsDef.COD03Z + "'"));
		DataTable tVar = _clsFun.FillTabSql("GesVariazioni", s, bPrimo: false, _strConSql);
		ClearColumnMaxLength(tVar);
		s = "SELECT * FROM TmpDiv WHERE tmp_tip='ETI' AND tmp_sta='S'";
		DataTable tTmp = _clsFun.FillTabSql("TmpDiv", s, bPrimo: false, _strConSql);
		HashSet<string> existingVarMap = new HashSet<string>();
		foreach (DataRow rVar in tVar.Rows)
		{
			string tip = rVar["var_tip"]?.ToString().Trim() ?? "";
			string num = rVar["var_num"]?.ToString().Trim() ?? "";
			string art = rVar["var_art"]?.ToString().Trim().PadLeft(7, '0') ?? "";
			existingVarMap.Add(tip + "|" + art + "|" + num);
		}
		foreach (DataRow y in tTmp.Rows)
		{
			string[] a = ((string)y["tmp_tmp"]).Split(';');
			if (a.Length > 4)
			{
				string sEan = (a.Length > 3 && !string.IsNullOrEmpty(a[3])) ? a[3].Trim() : "";
				string sArt = a[4].Trim();

				if (sArt == "SoloEan" || sArt == sEan || sArt.Length > 7)
				{
					if (string.IsNullOrEmpty(sEan) && sArt.Length > 7)
					{
						sEan = sArt;
					}
					if (!string.IsNullOrEmpty(sEan))
					{
						DataTable tArt = _clsQry.ArtSeek(sEan, "");
						if (tArt.Rows.Count > 0)
						{
							sArt = ((string)tArt.Rows[0]["tmp_art"]).Trim().PadLeft(7, '0');
						}
					}
				}
				else if (!string.IsNullOrEmpty(sArt))
				{
					sArt = sArt.PadLeft(7, '0');
				}

				if (!string.IsNullOrEmpty(sArt))
				{
					string etiKey = "ETI|" + sArt + "|000";
					if (existingVarMap.Add(etiKey))
					{
						DataRow x = tVar.NewRow();
						x["var_tip"] = "ETI";
						x["var_num"] = "000";
						x["var_ori"] = "ApPhone";
						x["var_art"] = sArt;
						x["var_dti"] = _clsFun.Str2Day(((string)y["tmp_fil"]).Substring(4, 8));
						x["var_ean"] = sEan;
						tVar.Rows.Add(x);
					}

					string posKey = "POS|" + sArt + "|000";
					if (existingVarMap.Add(posKey))
					{
						DataRow x2 = tVar.NewRow();
						x2["var_tip"] = "POS";
						x2["var_num"] = "000";
						x2["var_ori"] = "ApPhone";
						x2["var_art"] = sArt;
						x2["var_dti"] = _clsFun.Str2Day(((string)y["tmp_fil"]).Substring(4, 8));
						x2["var_ean"] = sEan;
						tVar.Rows.Add(x2);
					}
				}
			}
		}
		// Non cancellare o alterare TmpDiv all'avvio/ingresso form per non perdere le rilevazioni del terminale.
		progressBar1.Value = 0;
		progressBar1.Maximum = tVar.Rows.Count;
		progressBar1.Minimum = 0;
		dgv1.SuspendLayout();
		dgv3.SuspendLayout();
		Cursor.Current = Cursors.WaitCursor;
		HashSet<string> addedPos = new HashSet<string>();
		HashSet<string> addedEti = new HashSet<string>();
		LookupCache cache = new LookupCache();
		foreach (DataRow r in tRep.Rows)
		{
			cache.Reparti[r["tab_cod"].ToString()] = r;
		}
		foreach (DataRow r2 in tIva.Rows)
		{
			cache.Iva[r2["tab_cod"].ToString()] = r2;
		}
		foreach (DataRow r3 in tTgr.Rows)
		{
			cache.Tgr[r3["tab_cod"].ToString()] = r3;
		}
		foreach (DataRow r4 in tCal.Rows)
		{
			cache.Cal[r4["tab_cod"].ToString()] = r4;
		}
		foreach (DataRow r5 in tCat.Rows)
		{
			cache.Cat[r5["tab_cod"].ToString()] = r5;
		}
		foreach (DataRow r6 in tOri.Rows)
		{
			cache.Ori[r6["tab_cod"].ToString()] = r6;
		}
		if (dSet.Tables.Contains("GesOffArticoli"))
		{
			foreach (DataRow r7 in dSet.Tables["GesOffArticoli"].Rows)
			{
				cache.OffArt[r7["ofa_art"].ToString()] = r7;
			}
		}
		foreach (DataRow r8 in tOffEti.Rows)
		{
			cache.OffTes[r8["oft_cod"].ToString() + "|" + r8["ofa_art"].ToString()] = r8;
		}
		string sLisNeg = ((tLne != null && tLne.Rows.Count > 0) ? ((string)tLne.Rows[0]["tab_lis"]) : "");
		int totalVar = tVar.Rows.Count;
		progressBar1.Value = 0;
		progressBar1.Maximum = ((totalVar <= 0) ? 1 : totalVar);
		progressBar1.Minimum = 0;
		PreloadArticlesCache(tVar, sLisNeg);
		int idxVar = 0;
		foreach (DataRow y2 in tVar.Rows)
		{
			idxVar++;
			if (idxVar % 25 == 0 || idxVar == totalVar)
			{
				int pct = (int)((double)idxVar / (double)totalVar * 100.0);
				progressBar1.Value = Math.Min(idxVar, progressBar1.Maximum);
				string msgProgress = $"Elaborazione: {pct}% ({idxVar}/{totalVar})";
				lblVarCnt.Text = msgProgress;
				Text = "Gestione Variazioni - " + msgProgress;
				Application.DoEvents();
			}
			if ((string)y2["var_tip"] == _clsDef.VARPOS)
			{
				FillPos(dSet, y2, addedPos, cache);
			}
			else if ((string)y2["var_tip"] == _clsDef.VARETI)
			{
				if (strVarNum != "")
				{
					y2["var_inv"] = _clsDef.DIVDAD;
				}
				FillEti(dSet, tEti, y2, tOffEti, addedEti, cache);
			}
		}
		Text = "Gestione Variazioni";
		dSet.Tables.Clear();
		dSet.Clear();
		dSet.Dispose();
		dgv1.DataSource = tPos;
		dgv3.DataSource = tEti;
		dgv1.ResumeLayout();
		dgv3.ResumeLayout();
		Cursor.Current = Cursors.Default;
		lblVarCnt.Text = tPos.Rows.Count.ToString();
		DataView v = new DataView(tEti, "eti_inv='0'", "", DataViewRowState.CurrentRows);
		lblEtiCnt.Text = v.Count.ToString();
		FillRepCmbEti(tEti);
	}

	private void PreloadArticlesCache(DataTable tVar, string sLisNeg)
	{
		if (tVar == null || tVar.Rows.Count == 0)
		{
			return;
		}
		List<string> artCodes = new List<string>();
		foreach (DataRow r in tVar.Rows)
		{
			if (r["var_art"] != DBNull.Value)
			{
				string code = r["var_art"].ToString().Trim().PadLeft(7, '0');
				if (!string.IsNullOrEmpty(code) && !TryGetCachedArticleRows(code, out var _) && !artCodes.Contains(code))
				{
					artCodes.Add(code);
				}
			}
		}
		if (artCodes.Count == 0)
		{
			return;
		}
		if (_cacheMasterTable == null)
		{
			_cacheMasterTable = new clsGenTabTmp().TabTmpArt("ArtCod");
		}
		int totalChunks = (artCodes.Count + 499) / 500;
		int processedChunks = 0;
		string origTitle = Text;
		for (int i = 0; i < artCodes.Count; i += 500)
		{
			processedChunks++;
			int pct = (int)((double)processedChunks / (double)totalChunks * 100.0);
			int currentCount = Math.Min(i + 500, artCodes.Count);
			Text = $"Gestione Variazioni - Caricamento articoli {pct}% ({currentCount} / {artCodes.Count})...";
			if (lblVarCnt != null)
			{
				lblVarCnt.Text = $"Caricamento in corso: {pct}% ({currentCount}/{artCodes.Count})";
			}
			if (progressBar1 != null)
			{
				progressBar1.Maximum = artCodes.Count;
				progressBar1.Value = currentCount;
			}
			Application.DoEvents();
			List<string> chunk = artCodes.Skip(i).Take(500).ToList();
			string inClause = "'" + string.Join("','", chunk) + "'";
			string sSql = "SELECT ";
			sSql += "AnaArticoli.*, ";
			sSql += "AnaBarcode.ean_bil, AnaBarcode.ean_tip, AnaBarcode.ean_ean, AnaBarcode.ean_qta, AnaBarcode.ean_ecp, ";
			sSql += "AnaBarcode.ean_prv, AnaBarcode.ean_dti, AnaBarcode.ean_dtm, AnaBarcode.ean_pun, AnaBarcode.ean_ann ";
			sSql += "FROM AnaArticoli LEFT OUTER JOIN AnaBarcode ON AnaArticoli.art_cod = AnaBarcode.ean_art ";
			sSql = sSql + "WHERE art_cod IN (" + inClause + ") ";
			sSql += "ORDER BY AnaBarcode.ean_dti DESC";
			DataTable tBulkArt = _clsFun.FillTabSql("AnaArticoli", sSql, bPrimo: false, _strConSql);
			sSql = "SELECT liv_art, liv_lis, liv_prv, liv_dti, liv_dtf, liv_ann FROM (";
			sSql += "SELECT ROW_NUMBER() OVER (PARTITION BY liv_art, liv_lis ORDER BY liv_art, liv_dti DESC) AS ROW, ";
			sSql += "liv_art, liv_lis, liv_prv, liv_dti, liv_dtf, liv_ann ";
			sSql += "FROM GesLisVendita ";
			sSql = sSql + "WHERE liv_art IN (" + inClause + ") AND liv_dti <= " + _clsFun.DaySql(DateTime.Today);
			sSql += ") AS A WHERE ROW = 1";
			DataTable tBulkPrice = _clsFun.FillTabSql("GesLisVendita", sSql, bPrimo: false, _strConSql);
			Dictionary<string, List<DataRow>> priceDict = new Dictionary<string, List<DataRow>>();
			if (tBulkPrice != null)
			{
				foreach (DataRow pr in tBulkPrice.Rows)
				{
					string art = pr["liv_art"].ToString().Trim().PadLeft(7, '0');
					if (!priceDict.ContainsKey(art))
					{
						priceDict[art] = new List<DataRow>();
					}
					priceDict[art].Add(pr);
				}
			}
			sSql = "SELECT lia_art, lia_cos, lia_for, lia_arf, for_des FROM (";
			sSql += "SELECT ROW_NUMBER() OVER (PARTITION BY lia_art ORDER BY lia_dti DESC) AS ROW, ";
			sSql += "lia_art, lia_cos, lia_for, lia_arf, for_des ";
			sSql += "FROM GesLisAcquisto LEFT JOIN AnaFornitori ON AnaFornitori.for_cod = GesLisAcquisto.lia_for ";
			sSql = sSql + "WHERE lia_art IN (" + inClause + ") AND (lia_ann=0 OR lia_ann IS NULL) AND lia_dti <= " + _clsFun.DaySql(DateTime.Today);
			sSql += ") AS B WHERE ROW = 1";
			DataTable tBulkCost = _clsFun.FillTabSql("GesLisAcquisto", sSql, bPrimo: false, _strConSql);
			Dictionary<string, DataRow> costDict = new Dictionary<string, DataRow>();
			if (tBulkCost != null)
			{
				foreach (DataRow cr in tBulkCost.Rows)
				{
					costDict[cr["lia_art"].ToString().Trim().PadLeft(7, '0')] = cr;
				}
			}
			if (tBulkArt != null)
			{
				Func<DataRow, string, object, object> safeGet = (DataRow row, string colName, object defaultVal) => (row.Table.Columns.Contains(colName) && !DBNull.Value.Equals(row[colName])) ? row[colName] : defaultVal;
				Func<DataRow, bool> isNotAnn = delegate(DataRow p)
				{
					if (!p.Table.Columns.Contains("liv_ann") || DBNull.Value.Equals(p["liv_ann"]))
					{
						return true;
					}
					if (p["liv_ann"] is bool flag)
					{
						return !flag;
					}
					if (p["liv_ann"] is int num)
					{
						return num == 0;
					}
					if (p["liv_ann"] is byte b)
					{
						return b == 0;
					}
					if (p["liv_ann"] is short num2)
					{
						return num2 == 0;
					}
					string text = p["liv_ann"].ToString().Trim();
					return text == "0" || text == "" || text.ToLower() == "false";
				};
				Func<DataRow, string, bool> isLis = (DataRow p, string lisCode) => p.Table.Columns.Contains("liv_lis") && !DBNull.Value.Equals(p["liv_lis"]) && p["liv_lis"].ToString().Trim() == lisCode;
				IEnumerable<IGrouping<string, DataRow>> artGroups = from dataRow2 in tBulkArt.AsEnumerable()
					group dataRow2 by dataRow2["art_cod"].ToString().Trim().PadLeft(7, '0');
				HashSet<string> loadedCodes = new HashSet<string>();
				foreach (IGrouping<string, DataRow> grp in artGroups)
				{
					string artCod = grp.Key.Trim().PadLeft(7, '0');
					List<DataRow> articleRows = new List<DataRow>();
					foreach (DataRow y in grp)
					{
						DataRow x = _cacheMasterTable.NewRow();
						x["tmp_art"] = safeGet(y, "art_cod", "");
						x["tmp_ard"] = safeGet(y, "art_des", "");
						x["tmp_arb"] = safeGet(y, "art_deb", "");
						x["tmp_sta"] = safeGet(y, "art_sta", "");
						x["tmp_rep"] = safeGet(y, "art_rep", "");
						string ec1 = safeGet(y, "art_ec1", "").ToString();
						string ec2 = safeGet(y, "art_ec2", "").ToString();
						string ec3 = safeGet(y, "art_ec3", "").ToString();
						x["tmp_ecr"] = (ec1 + new string(' ', 3)).Substring(0, 3);
						DataRow dataRow = x;
						dataRow["tmp_ecr"] = dataRow["tmp_ecr"]?.ToString() + (ec2 + new string(' ', 3)).Substring(0, 3);
						dataRow = x;
						dataRow["tmp_ecr"] = dataRow["tmp_ecr"]?.ToString() + (ec3 + new string(' ', 3)).Substring(0, 3);
						x["tmp_iva"] = safeGet(y, "art_iva", "");
						x["tmp_umi"] = safeGet(y, "art_umi", "");
						x["tmp_pxc"] = safeGet(y, "art_pxc", 0);
						x["tmp_tgr"] = safeGet(y, "art_tgr", "");
						x["tmp_pne"] = safeGet(y, "art_pne", 0);
						x["tmp_ori"] = safeGet(y, "art_ori", "");
						x["tmp_cal"] = safeGet(y, "art_cal", "");
						x["tmp_cat"] = safeGet(y, "art_cat", "");
						x["tmp_tar"] = safeGet(y, "art_tar", "");
						x["tmp_sfr"] = safeGet(y, "art_sfr", "");
						x["tmp_reb"] = safeGet(y, "art_reb", 0);
						x["tmp_gsc"] = safeGet(y, "art_gsc", 0);
						x["tmp_plu"] = safeGet(y, "art_plu", "").ToString().Trim();
						x["tmp_tas"] = safeGet(y, "art_tas", "").ToString().Trim();
						x["tmp_bpz"] = safeGet(y, "art_bpz", 0);
						x["tmp_tra"] = safeGet(y, "art_tra", "");
						x["tmp_eti"] = safeGet(y, "art_eti", "");
						x["tmp_img"] = safeGet(y, "art_img", "");
						x["tmp_bil"] = safeGet(y, "ean_bil", false);
						x["tmp_ean"] = safeGet(y, "ean_ean", "");
						x["tmp_eaq"] = safeGet(y, "ean_qta", 0);
						x["tmp_eap"] = safeGet(y, "ean_prv", 0);
						x["tmp_ecp"] = safeGet(y, "ean_ecp", "");
						x["tmp_eaa"] = safeGet(y, "ean_ann", false);
						x["tmp_epu"] = safeGet(y, "ean_pun", 0);
						if (priceDict.ContainsKey(artCod))
						{
							List<DataRow> prices = priceDict[artCod];
							DataRow pPos = prices.FirstOrDefault((DataRow p) => isLis(p, _clsDef.LISPOS) && isNotAnn(p));
							DataRow pNeg = ((!string.IsNullOrEmpty(sLisNeg)) ? prices.FirstOrDefault((DataRow p) => isLis(p, sLisNeg) && isNotAnn(p)) : null);
							DataRow pFor = prices.FirstOrDefault((DataRow p) => isLis(p, _clsDef.LISFOR) && isNotAnn(p));
							decimal prv = default(decimal);
							string lis = "";
							if (pNeg != null && pNeg.Table.Columns.Contains("liv_prv") && !DBNull.Value.Equals(pNeg["liv_prv"]))
							{
								prv = Convert.ToDecimal(pNeg["liv_prv"]);
								lis = sLisNeg;
							}
							else if (pPos != null && pPos.Table.Columns.Contains("liv_prv") && !DBNull.Value.Equals(pPos["liv_prv"]))
							{
								prv = Convert.ToDecimal(pPos["liv_prv"]);
								lis = _clsDef.LISPOS;
							}
							else if (pFor != null && pFor.Table.Columns.Contains("liv_prv") && !DBNull.Value.Equals(pFor["liv_prv"]))
							{
								prv = Convert.ToDecimal(pFor["liv_prv"]);
								lis = _clsDef.LISFOR;
							}
							x["tmp_prv"] = prv;
							x["tmp_lis"] = lis;
							decimal ppv = default(decimal);
							if (pPos != null && pPos.Table.Columns.Contains("liv_prv") && !DBNull.Value.Equals(pPos["liv_prv"]))
							{
								ppv = Convert.ToDecimal(pPos["liv_prv"]);
							}
							else if (pFor != null && pFor.Table.Columns.Contains("liv_prv") && !DBNull.Value.Equals(pFor["liv_prv"]))
							{
								ppv = Convert.ToDecimal(pFor["liv_prv"]);
							}
							x["tmp_ppv"] = ppv;
						}
						if (costDict.ContainsKey(artCod))
						{
							DataRow cRow = costDict[artCod];
							x["tmp_for"] = cRow["lia_for"];
							x["tmp_fod"] = cRow["for_des"];
							x["tmp_arf"] = cRow["lia_arf"];
							x["tmp_cos"] = cRow["lia_cos"];
						}
						_cacheMasterTable.Rows.Add(x);
						articleRows.Add(x);
					}
					DataRow[] rowsArr = articleRows.ToArray();
					CacheArticleRows(artCod, rowsArr);
					loadedCodes.Add(artCod);
					loadedCodes.Add(artCod.Trim());
					loadedCodes.Add(artCod.Trim().PadLeft(7, '0'));
				}
				foreach (string reqCode in chunk)
				{
					string reqTrimmed = reqCode.Trim();
					string reqPadded = reqTrimmed.PadLeft(7, '0');
					if (!loadedCodes.Contains(reqCode) && !loadedCodes.Contains(reqTrimmed) && !loadedCodes.Contains(reqPadded))
					{
						CacheArticleRows(reqCode, new DataRow[0]);
					}
				}
			}
			if (tBulkArt != null)
			{
				tBulkArt.Dispose();
				tBulkArt = null;
			}
			if (tBulkPrice != null)
			{
				tBulkPrice.Dispose();
				tBulkPrice = null;
			}
			if (tBulkCost != null)
			{
				tBulkCost.Dispose();
				tBulkCost = null;
			}
			priceDict.Clear();
			costDict.Clear();
		}
		Text = origTitle;
	}

	private void FillRepCmbEti(DataTable tabEti)
	{
		if (tabEti.Rows.Count <= 0)
		{
			return;
		}
		string sRep = "";
		if (cmbEtiRep.SelectedValue != null && cmbEtiRep.SelectedValue.ToString() != "")
		{
			sRep = cmbEtiRep.SelectedValue.ToString();
		}
		string s = "SELECT tab_cod, tab_des FROM TabReparti ";
		DataTable tRep = _clsFun.FillTabSql("TabReparti", s, bPrimo: false, _strConSql);
		DataTable t = tRep.Clone();
		DataRow x = t.NewRow();
		x["tab_cod"] = "";
		x["tab_des"] = " Tutti";
		t.Rows.Add(x);
		HashSet<string> existingCodes = new HashSet<string>();
		existingCodes.Add("");
		Dictionary<string, string> repDescs = new Dictionary<string, string>();
		if (tRep != null)
		{
			foreach (DataRow r in tRep.Rows)
			{
				if (r["tab_cod"] != DBNull.Value)
				{
					repDescs[r["tab_cod"].ToString().Trim()] = (r["tab_des"] != DBNull.Value) ? r["tab_des"].ToString() : "";
				}
			}
		}
		foreach (DataRow y in tabEti.Rows)
		{
			if (y["eti_rep"] != DBNull.Value)
			{
				string repCod = y["eti_rep"].ToString().Trim();
				if (existingCodes.Add(repCod))
				{
					x = t.NewRow();
					x["tab_cod"] = repCod;
					x["tab_des"] = repDescs.TryGetValue(repCod, out string des) ? des : "";
					t.Rows.Add(x);
				}
			}
		}
		cmbEtiRep.DataSource = t;
		cmbEtiRep.DisplayMember = "tab_des";
		cmbEtiRep.ValueMember = "tab_cod";
		cmbEtiRep.SelectedValue = sRep;
	}

	private void FillPos(DataSet dasSet, DataRow rowVar)
	{
		FillPos(dasSet, rowVar, new HashSet<string>(), new LookupCache());
	}

	private void FillPos(DataSet dasSet, DataRow rowVar, HashSet<string> addedPos, LookupCache cache)
	{
		string sArt = ((rowVar["var_art"] != DBNull.Value) ? rowVar["var_art"].ToString().Trim().PadLeft(7, '0') : "");
		IEnumerable<DataRow> tArtRows = null;
		if (TryGetCachedArticleRows(sArt, out var cachedRows))
		{
			tArtRows = cachedRows;
		}
		else
		{
			string sLisNeg = "";
			if (dasSet != null && dasSet.Tables.Contains("TabLisNegozi") && dasSet.Tables["TabLisNegozi"].Rows.Count > 0)
			{
				sLisNeg = (string)dasSet.Tables["TabLisNegozi"].Rows[0]["tab_lis"];
			}
			DataTable tArt = _clsQry.SeekArtEan(sArt, "LNE" + sLisNeg);
			DataRow[] rows = ((tArt != null && tArt.Rows.Count > 0) ? tArt.Select() : new DataRow[0]);
			CacheArticleRows(sArt, rows);
			tArtRows = rows;
		}
		FillPosInternal(dasSet, rowVar, tArtRows, addedPos, cache);
	}

	private void FillPosInternal(DataSet dasSet, DataRow rowVar, IEnumerable<DataRow> tArtRows, HashSet<string> addedPos, LookupCache cache)
	{
		string s = "";
		string sLisNeg = "";
		if (dasSet.Tables.Contains("TabLisNegozi") && dasSet.Tables["TabLisNegozi"].Rows.Count > 0)
		{
			sLisNeg = (string)dasSet.Tables["TabLisNegozi"].Rows[0]["tab_lis"];
		}
		s = "LNE" + sLisNeg;
		if ((string)rowVar["var_art"] == "0002382")
		{
			Console.WriteLine("aaaa");
		}
		if (tArtRows == null)
		{
			return;
		}
		foreach (DataRow k in tArtRows)
		{
			string key = (string)k["tmp_art"] + "|" + (string)k["tmp_ean"];
			if (addedPos.Contains(key))
			{
				continue;
			}
			addedPos.Add(key);
			DataRow x = dasSet.Tables["PosArt"].NewRow();
			x["pos_inv"] = _clsDef.DIVDAD;
			if (!DBNull.Value.Equals(rowVar["var_inv"]) && ((string)rowVar["var_inv"]).Trim() != "")
			{
				x["pos_inv"] = rowVar["var_inv"];
			}
			x["pos_art"] = k["tmp_art"];
			x["pos_ard"] = k["tmp_ard"];
			x["pos_arb"] = k["tmp_arb"];
			if (((string)x["pos_arb"]).Trim() == "")
			{
				x["pos_arb"] = x["pos_ard"];
			}
			if (((string)x["pos_arb"]).Length > 20)
			{
				x["pos_arb"] = ((string)x["pos_arb"]).Substring(0, 20);
			}
			x["pos_sta"] = k["tmp_sta"];
			x["pos_iva"] = k["tmp_iva"];
			x["pos_umi"] = k["tmp_umi"];
			x["pos_tgr"] = k["tmp_tgr"];
			x["pos_pxc"] = k["tmp_pxc"];
			decimal pnePos = (!DBNull.Value.Equals(k["tmp_pne"])) ? Convert.ToDecimal(k["tmp_pne"]) : 0m;
			decimal prvPos = (!DBNull.Value.Equals(k["tmp_prv"])) ? Convert.ToDecimal(k["tmp_prv"]) : 0m;
			decimal tgvPos = 1m;
			if (cache != null && cache.Tgr != null && k["tmp_tgr"] != DBNull.Value && cache.Tgr.ContainsKey(k["tmp_tgr"].ToString()))
			{
				tgvPos = Convert.ToDecimal(cache.Tgr[k["tmp_tgr"].ToString()]["tab_val"]);
			}
			if (pnePos > 500m || (prvPos > 0m && pnePos > 0m && tgvPos > 0m && (prvPos / (pnePos / tgvPos)) > 500m))
			{
				pnePos = 0m;
			}
			x["pos_pne"] = pnePos;
			x["pos_rep"] = k["tmp_rep"];
			x["pos_ecr"] = k["tmp_ecr"];
			if (_strPar039SoloPrezziNeg2Pos == "S" && (string)k["tmp_lis"] == _clsDef.LISFOR)
			{
				x["pos_prv"] = 0;
				x["pos_msg"] = "Prezzo negozio mancante ";
			}
			else
			{
				x["pos_prv"] = k["tmp_prv"];
			}
			x["pos_ean"] = k["tmp_ean"];
			if (string.IsNullOrEmpty(Convert.ToString(x["pos_ean"]).Trim()))
			{
				if (!DBNull.Value.Equals(k["tmp_plu"]) && !string.IsNullOrEmpty(Convert.ToString(k["tmp_plu"]).Trim()))
					x["pos_ean"] = Convert.ToString(k["tmp_plu"]).Trim();
				else if (!DBNull.Value.Equals(k["tmp_art"]) && !string.IsNullOrEmpty(Convert.ToString(k["tmp_art"]).Trim()))
					x["pos_ean"] = Convert.ToString(k["tmp_art"]).Trim();
			}
			x["pos_eaq"] = k["tmp_eaq"];
			x["pos_eap"] = k["tmp_eap"];
			x["pos_epu"] = k["tmp_epu"];
			x["pos_ecp"] = k["tmp_ecp"];
			x["pos_eaa"] = k["tmp_eaa"];
			x["pos_ori"] = rowVar["var_ori"];
			x["pos_day"] = rowVar["var_dti"];
			if ((decimal)x["pos_eap"] > 0m)
			{
				x["pos_prv"] = (decimal)x["pos_eap"];
			}
			x["bil_bil"] = k["tmp_bil"];
			x["bil_reb"] = k["tmp_reb"];
			x["bil_gsc"] = k["tmp_gsc"];
			x["bil_plu"] = k["tmp_plu"];
			x["bil_bpz"] = k["tmp_bpz"];
			x["bil_bpz"] = k["tmp_bpz"];
			x["bil_tar"] = k["tmp_tar"];
			x["bil_tra"] = k["tmp_tra"];
			x["bil_tas"] = k["tmp_tas"];
			x["bil_img"] = k["tmp_img"];
			if ((string)x["pos_sta"] == _clsDef.STAREP)
			{
				if (((string)x["pos_rep"]).Trim() == "" || Convert.ToInt16(x["pos_rep"]) == 0)
				{
					DataRow dataRow = x;
					dataRow["pos_msg"] = dataRow["pos_msg"]?.ToString() + "Reparto mancante";
				}
			}
			else
			{
				if ((string)x["pos_sta"] != _clsDef.STAREP && (string)x["pos_sta"] != _clsDef.STAPRE && Convert.ToDecimal(x["pos_prv"]) == 0m)
				{
					DataRow dataRow = x;
					dataRow["pos_msg"] = dataRow["pos_msg"]?.ToString() + "Prezzo mancante ";
				}
				if (((string)x["pos_arb"]).Trim() == "")
				{
					DataRow dataRow = x;
					dataRow["pos_msg"] = dataRow["pos_msg"]?.ToString() + "Descriz. mancante ";
				}
				if ((string)x["pos_sta"] != _clsDef.STAREP && ((string)x["pos_ean"]).Trim() == "")
				{
					DataRow dataRow = x;
					dataRow["pos_msg"] = dataRow["pos_msg"]?.ToString() + "Ean mancante ";
				}
				if (((string)x["pos_iva"]).Trim() == "")
				{
					DataRow dataRow = x;
					dataRow["pos_msg"] = dataRow["pos_msg"]?.ToString() + "IVA mancante";
				}
				else if (!cache.Iva.ContainsKey(x["pos_iva"].ToString()))
				{
					x["pos_msg"] = "IVA " + x["pos_iva"]?.ToString() + " non presente in cassa ";
				}
				if (((string)x["pos_rep"]).Trim() == "" || Convert.ToInt16(x["pos_rep"]) == 0)
				{
					DataRow dataRow = x;
					dataRow["pos_msg"] = dataRow["pos_msg"]?.ToString() + "Reparto mancante";
				}
				else if (!cache.Reparti.ContainsKey(x["pos_rep"].ToString()))
				{
					x["pos_msg"] = "Reparto " + x["pos_rep"]?.ToString() + " non presente in cassa ";
				}
				if ((bool)x["bil_bil"] && ((string)x["pos_ean"]).Length != 13 && (string)x["pos_sta"] != _clsDef.STAREP && ((string)x["pos_ean"]).Trim().Length > 7)
				{
					DataRow dataRow = x;
					dataRow["pos_msg"] = dataRow["pos_msg"]?.ToString() + "Barcode < 13 ";
				}
				if (((string)x["pos_ean"]).Trim().Length > 7)
				{
					s = ((string)x["pos_ean"]).Trim();
					if (s.Substring(0, 1) == "2" && s.Length == 13 && s.Substring(7, 6) == "000000")
					{
						Console.WriteLine("Ean a peso");
					}
					else
					{
						try
						{
							if (s.Length > 1)
							{
								s = s.Substring(0, s.Length - 1) + new clsCtrlCodici().FindMod10Digit(s.Substring(0, s.Length - 1));
							}
						}
						catch
						{
						}
						if (((string)x["pos_ean"]).Trim() != s)
						{
							DataRow dataRow = x;
							dataRow["pos_msg"] = dataRow["pos_msg"]?.ToString() ?? "";
						}
					}
					s = ((string)x["pos_ean"]).Trim();
					if ((bool)x["bil_bil"] && (s.Substring(0, 1) != "2" || s.Length != 13 || s.Substring(7, 6) != "000000"))
					{
						DataRow dataRow = x;
						dataRow["pos_msg"] = dataRow["pos_msg"]?.ToString() + "Barcode bilancia non corretto";
					}
				}
			}
			if (((string)x["pos_msg"]).Trim() != "")
			{
				if (((string)x["pos_msg"]).Trim() == "Prezzo mancante")
				{
					if (_strCnfPar.Length == 0 || _strCnfPar.Substring(0, 1) != "S")
					{
						x["pos_inv"] = _clsDef.DIVSOS;
					}
				}
				else
				{
					x["pos_inv"] = _clsDef.DIVSOS;
				}
			}
			if (cache.OffArt.ContainsKey(k["tmp_art"].ToString()))
			{
				FillOffArt(dasSet.Tables["TabOfferte"], x, cache.OffArt[k["tmp_art"].ToString()]);
			}
			dasSet.Tables["PosArt"].Rows.Add(x);
		}
	}

	private void FillEti(DataSet dasSet, DataTable tabEti, DataRow rowVar, DataTable tabOff)
	{
		FillEti(dasSet, tabEti, rowVar, tabOff, new HashSet<string>(), new LookupCache());
	}

	private void FillEti(DataSet dasSet, DataTable tabEti, DataRow rowVar, DataTable tabOff, HashSet<string> addedEti, LookupCache cache)
	{
		string sArt = ((rowVar["var_art"] != DBNull.Value) ? rowVar["var_art"].ToString().Trim().PadLeft(7, '0') : "");
		IEnumerable<DataRow> tArtRows = null;
		if (TryGetCachedArticleRows(sArt, out var cachedRows))
		{
			tArtRows = cachedRows;
		}
		else
		{
			string sLisNeg = "";
			if (dasSet != null && dasSet.Tables.Contains("TabLisNegozi") && dasSet.Tables["TabLisNegozi"].Rows.Count > 0)
			{
				sLisNeg = (string)dasSet.Tables["TabLisNegozi"].Rows[0]["tab_lis"];
			}
			DataTable tArt = _clsQry.SeekArtEan(sArt, "LNE" + sLisNeg);
			DataRow[] rows = ((tArt != null && tArt.Rows.Count > 0) ? tArt.Select() : new DataRow[0]);
			CacheArticleRows(sArt, rows);
			tArtRows = rows;
		}
		FillEtiInternal(dasSet, tabEti, rowVar, tabOff, tArtRows, addedEti, cache);
	}

	private void FillEtiInternal(DataSet dasSet, DataTable tabEti, DataRow rowVar, DataTable tabOff, IEnumerable<DataRow> tArtRows, HashSet<string> addedEti, LookupCache cache)
	{
		if (tabEti == null)
		{
			return;
		}
		if (!tabEti.Columns.Contains("eti_inv"))
		{
			DataTable tmplEti = new clsGenTabTmp().TabTmpEti("EtiArt");
			foreach (DataColumn col in tmplEti.Columns)
			{
				if (!tabEti.Columns.Contains(col.ColumnName))
				{
					tabEti.Columns.Add(new DataColumn(col.ColumnName, col.DataType)
					{
						Caption = col.Caption,
						DefaultValue = col.DefaultValue
					});
				}
			}
		}
		string s = "";
		string sLisNeg = "";
		if (dasSet.Tables.Contains("TabLisNegozi") && dasSet.Tables["TabLisNegozi"].Rows.Count > 0)
		{
			sLisNeg = (string)dasSet.Tables["TabLisNegozi"].Rows[0]["tab_lis"];
		}
		s = "LNE" + sLisNeg;
		if (tArtRows == null)
		{
			return;
		}
		foreach (DataRow k in tArtRows)
		{
			if ((bool)k["tmp_eaa"])
			{
				continue;
			}
			string key = k["tmp_art"]?.ToString();
			if (key == null || addedEti.Contains(key))
			{
				continue;
			}
			addedEti.Add(key);
			DataRow x = tabEti.NewRow();
			x["eti_inv"] = _clsDef.DIVDAD;
			if (rowVar != null && rowVar.Table.Columns.Contains("var_inv") && !DBNull.Value.Equals(rowVar["var_inv"]) && ((string)rowVar["var_inv"]).Trim() != "")
			{
				x["eti_inv"] = rowVar["var_inv"];
			}
			x["eti_qta"] = 1;
			if (rowVar != null && rowVar.Table.Columns.Contains("var_etq") && !DBNull.Value.Equals(rowVar["var_etq"]) && (decimal)rowVar["var_etq"] > 0m)
			{
				x["eti_qta"] = rowVar["var_etq"];
			}
			x["eti_art"] = k["tmp_art"];
			x["eti_ard"] = k["tmp_ard"];
			x["eti_sta"] = k["tmp_sta"];
			x["eti_iva"] = k["tmp_iva"];
			x["eti_umi"] = k["tmp_umi"];
			x["eti_tgr"] = k["tmp_tgr"];
			x["eti_pxc"] = k["tmp_pxc"];
			x["eti_plu"] = k["tmp_plu"];
			x["eti_rep"] = k["tmp_rep"];
			x["eti_ori"] = ((rowVar != null && rowVar.Table.Columns.Contains("var_ori")) ? rowVar["var_ori"] : "");
			x["eti_tgv"] = 1;
			if (cache != null && cache.Tgr != null && k["tmp_tgr"] != DBNull.Value && cache.Tgr.ContainsKey(k["tmp_tgr"].ToString()))
			{
				x["eti_tgv"] = cache.Tgr[k["tmp_tgr"].ToString()]["tab_val"];
			}
			else if (dasSet != null && dasSet.Tables.Contains("TabTipoGrammatura") && dasSet.Tables["TabTipoGrammatura"] != null)
			{
				DataRow[] j = dasSet.Tables["TabTipoGrammatura"].Select("tab_cod='" + (string)k["tmp_tgr"] + "'");
				if (j.Length != 0)
				{
					x["eti_tgv"] = j[0]["tab_val"];
				}
			}
			decimal pneEti = (!DBNull.Value.Equals(k["tmp_pne"])) ? Convert.ToDecimal(k["tmp_pne"]) : 0m;
			decimal prvEti = (!DBNull.Value.Equals(k["tmp_prv"])) ? Convert.ToDecimal(k["tmp_prv"]) : 0m;
			decimal tgvEti = (!DBNull.Value.Equals(x["eti_tgv"])) ? Convert.ToDecimal(x["eti_tgv"]) : 1m;
			if (pneEti > 500m || (prvEti > 0m && pneEti > 0m && tgvEti > 0m && (prvEti / (pneEti / tgvEti)) > 500m))
			{
				pneEti = 0m;
			}
			x["eti_pne"] = pneEti;
			x["eti_fod"] = k["tmp_fod"];
			x["eti_arf"] = k["tmp_arf"];
			x["eti_img"] = k["tmp_img"];
			if (cache != null && cache.Reparti != null && k["tmp_rep"] != DBNull.Value && cache.Reparti.ContainsKey(k["tmp_rep"].ToString()))
			{
				x["eti_red"] = (string)cache.Reparti[k["tmp_rep"].ToString()]["tab_des"];
			}
			if (_strPar039SoloPrezziNeg2Pos == "S" && (string)k["tmp_lis"] == _clsDef.LISFOR)
			{
				x["eti_prv"] = 0;
				x["eti_pve"] = 0;
				x["eti_msg"] = "Prezzo negozio mancante ";
			}
			else
			{
				x["eti_pve"] = k["tmp_prv"];
				x["eti_prv"] = k["tmp_prv"];
			}
			x["eti_eti"] = k["tmp_eti"];
			x["eti_ecr"] = k["tmp_ecr"];
			x["eti_tas"] = k["tmp_tas"];
			bool bHasVarOff = rowVar != null && rowVar.Table.Columns.Contains("var_off") && !DBNull.Value.Equals(rowVar["var_off"]);
			if (bHasVarOff && ((string)rowVar["var_off"]).Trim() == "PRO")
			{
				if (dasSet != null && dasSet.Tables.Contains("GesLisVendita") && dasSet.Tables["GesLisVendita"] != null && dasSet.Tables["GesLisVendita"].Rows.Count > 0)
				{
					DataRow[] j = dasSet.Tables["GesLisVendita"].Select("liv_art='" + (string)k["tmp_art"] + "' AND liv_dti<=" + _clsFun.DayMdb(DateTime.Today));
					if (j.Length != 0)
					{
						x["eti_off"] = "";
						x["eti_oft"] = "PRO";
						x["eti_pve"] = k["tmp_ppv"];
						x["eti_odi"] = (DateTime)j[0]["liv_dti"];
						x["eti_odf"] = (DateTime)j[0]["liv_dtf"];
						x["eti_cam"] = "";
					}
				}
			}
			else if (bHasVarOff && ((string)rowVar["var_off"]).Trim() != "")
			{
				string offKey = rowVar["var_off"].ToString() + "|" + k["tmp_art"].ToString();
				if (cache != null && cache.OffTes != null && cache.OffTes.ContainsKey(offKey))
				{
					DataRow offRow = cache.OffTes[offKey];
					x["eti_off"] = (string)offRow["ofa_cod"];
					x["eti_oft"] = (string)offRow["ofa_tip"];
					x["eti_prv"] = (decimal)offRow["ofa_val"];
					x["eti_odi"] = (DateTime)offRow["oft_dti"];
					x["eti_odf"] = (DateTime)offRow["oft_dtf"];
					x["eti_cam"] = (string)offRow["oft_cam"];
					x["eti_xem"] = 0;
					if (!DBNull.Value.Equals(offRow["ofa_xem"]))
					{
						x["eti_xem"] = (decimal)offRow["ofa_xem"];
					}
					x["eti_xen"] = 0;
					if (!DBNull.Value.Equals(offRow["ofa_xen"]))
					{
						x["eti_xen"] = (decimal)offRow["ofa_xen"];
					}
					if ((string)x["eti_oft"] == _clsDef.OFASCO)
					{
						x["eti_prv"] = (decimal)x["eti_pve"];
						x["eti_pve"] = (decimal)offRow["ofa_val"];
						DataRow dataRow = x;
						dataRow["eti_ori"] = dataRow["eti_ori"]?.ToString() + " - Sc." + (decimal)x["eti_pve"] + "%";
					}
				}
			}
			if (rowVar != null && rowVar.Table.Columns.Contains("var_ean") && !DBNull.Value.Equals(rowVar["var_ean"]) && ((string)rowVar["var_ean"]).Trim() != "")
			{
				x["eti_ean"] = rowVar["var_ean"];
			}
			else
			{
				x["eti_ean"] = k["tmp_ean"];
			}
			if (rowVar != null && rowVar.Table.Columns.Contains("var_dti") && !DBNull.Value.Equals(rowVar["var_dti"]))
			{
				x["eti_day"] = rowVar["var_dti"];
			}
			else
			{
				x["eti_day"] = DateTime.Today;
			}
			if (Convert.ToDecimal(x["eti_prv"]) == 0m)
			{
				DataRow dataRow = x;
				dataRow["eti_msg"] = dataRow["eti_msg"]?.ToString() + " Prezzo mancante ";
			}
			if (((string)x["eti_msg"]).Trim() != "")
			{
				x["eti_inv"] = _clsDef.DIVSOS;
			}
			if (((string)k["tmp_ori"]).Trim() != "" && cache.Ori.ContainsKey(k["tmp_ori"].ToString()))
			{
				x["eti_bor"] = (string)cache.Ori[k["tmp_ori"].ToString()]["tab_des"];
			}
			if (((string)k["tmp_cal"]).Trim() != "" && cache.Cal.ContainsKey(k["tmp_cal"].ToString()))
			{
				x["eti_bcl"] = (string)cache.Cal[k["tmp_cal"].ToString()]["tab_des"];
			}
			if (((string)k["tmp_cat"]).Trim() != "" && cache.Cat.ContainsKey(k["tmp_cat"].ToString()))
			{
				x["eti_bct"] = (string)cache.Cat[k["tmp_cat"].ToString()]["tab_des"];
			}
			tabEti.Rows.Add(x);
		}
	}

	private void FillOfferte()
	{
		string s = "SELECT tab_cod, tab_des, tab_art, tab_off FROM TabStato";
		DataTable tSta = _clsFun.FillTabSql("TabStato", s, bPrimo: false, _strConSql);
		s = "SELECT ";
		s += "GesOffTestate.* ";
		s += "FROM GesOffTestate ";
		s += "WHERE ";
		s = s + "GesOffTestate.oft_sta <> '" + _clsDef.STACLO + "' AND GesOffTestate.oft_sta <> '" + _clsDef.STACAN + "' ";
		s += "ORDER BY oft_dti DESC";
		DataTable t = _clsFun.FillTabSql("GesOffTestate", s, bPrimo: false, _strConSql);
		t.Columns.Add(new DataColumn
		{
			DataType = Type.GetType("System.String"),
			ColumnName = "OftStd",
			Caption = "Stato",
			MaxLength = 20,
			ReadOnly = false,
			DefaultValue = ""
		});
		t.DefaultView.AllowDelete = false;
		t.DefaultView.AllowEdit = false;
		t.DefaultView.AllowNew = false;
		foreach (DataRow y in t.Rows)
		{
			if ((string)y["oft_cod"] == "008")
			{
				Console.WriteLine("aaaa");
			}
			if ((string)y["oft_sta"] == _clsDef.STANOA || (string)y["oft_sta"] == _clsDef.STADAA)
			{
				DateTime dtI = (y["oft_dti"] != DBNull.Value) ? Convert.ToDateTime(y["oft_dti"]).Date : DateTime.MaxValue;
				DateTime dtF = (y["oft_dtf"] != DBNull.Value) ? Convert.ToDateTime(y["oft_dtf"]).Date : DateTime.MinValue;
				if (dtI <= DateTime.Today && dtF >= DateTime.Today)
				{
					y["oft_sta"] = _clsDef.STADAA;
				}
				else
				{
					y["oft_sta"] = _clsDef.STANOA;
				}
				if ((string)y["oft_sta"] == _clsDef.STADAA)
				{
					FillVarOffArt((string)y["oft_cod"], (string)y["oft_sta"], (string)y["oft_yea"]);
				}
			}
			else if ((string)y["oft_sta"] == _clsDef.STAATT || (string)y["oft_sta"] == _clsDef.STADAC)
			{
				DateTime dtF = (y["oft_dtf"] != DBNull.Value) ? Convert.ToDateTime(y["oft_dtf"]).Date : DateTime.MaxValue;
				if (dtF < DateTime.Today)
				{
					y["oft_sta"] = _clsDef.STADAC;
				}
				if ((string)y["oft_sta"] == _clsDef.STADAC)
				{
					FillVarOffArt((string)y["oft_cod"], (string)y["oft_sta"], (string)y["oft_yea"]);
				}
			}
			DataRow[] j = tSta.Select("tab_cod='" + y["oft_sta"]?.ToString() + "'");
			if (j.Length != 0)
			{
				y["OftStd"] = j[0]["tab_des"];
			}
		}
		dgv2.SuspendLayout();
		dgv2.DataSource = t;
		dgv2.ResumeLayout();
	}

	private void FillLisPromo()
	{
		string s = "";
		DataTable tLiv = (DataTable)dgv3.DataSource;
		DateTime dDay = DateTime.Today;
		DataTable t = _clsQry.LivPromo();
		DataTable tPro = t.Clone();
		foreach (DataRow y in t.Rows)
		{
			if ((string)y["liv_sta"] == _clsDef.STADAA && (DateTime)y["liv_dtf"] >= dDay)
			{
				tPro.ImportRow(y);
			}
			else if ((string)y["liv_sta"] == _clsDef.STAATT && (DateTime)y["liv_dtf"] < dDay)
			{
				tPro.ImportRow(y);
			}
		}
		foreach (DataRow y2 in tPro.Rows)
		{
			_clsVar.Variazioni((string)y2["liv_art"], "Forza", "Promo", _clsDef.VARALL);
			s = _clsDef.STAATT;
			if ((DateTime)y2["liv_dti"] < DateTime.Today)
			{
				s = _clsDef.STACLO;
			}
			s = "UPDATE GesLisVendita SET liv_sta='" + s + "' WHERE ";
			s = s + "liv_lis='" + _clsDef.LISPRO + "' AND ";
			s = s + "liv_art='" + y2["liv_art"]?.ToString() + "' AND ";
			s = s + "liv_dti=" + _clsFun.DaySql((DateTime)y2["liv_dti"]) + " AND ";
			s = s + "liv_dtf=" + _clsFun.DaySql((DateTime)y2["liv_dtf"]);
			_clsFun.SqlWrite(s, _strConSql);
		}
		dgv3.Invalidate();
	}

	private void FillVarOffArt(string strCod, string strSta, string strYea)
	{
		DataTable tPos = (DataTable)dgv1.DataSource;
		if (tPos == null) return;
		DataTable tEti = (DataTable)dgv3.DataSource;

		string s = "SELECT * FROM GesVariazioni WHERE var_num='" + _clsDef.COD04X + "'";
		DataTable tVar = _clsFun.FillTabSql("GesVariazioni", s, bPrimo: false, _strConSql);
		s = "SELECT * FROM TabOfferte WHERE tab_ann=0";
		DataTable tTof = _clsFun.FillTabSql("TabOfferte", s, bPrimo: false, _strConSql);
		s = "SELECT * FROM TabReparti WHERE tab_ann=0";
		DataTable tRep = _clsFun.FillTabSql("TabReparti", s, bPrimo: false, _strConSql);
		s = "SELECT * FROM TabIva WHERE tab_ann=0";
		DataTable tIva = _clsFun.FillTabSql("TabIva", s, bPrimo: false, _strConSql);
		DataTable tLne = _clsFun.FillTabSql("TabLisNegozi", "SELECT tab_lis FROM TabNegozi WHERE tab_tip = 'L'", bPrimo: false, _strConSql);

		DataSet dSet = new DataSet();
		try
		{
			Action<DataTable> safeAdd = tbl =>
			{
				if (tbl != null && !dSet.Tables.Contains(tbl.TableName))
				{
					if (tbl.DataSet != null) tbl.DataSet.Tables.Remove(tbl);
					dSet.Tables.Add(tbl);
				}
			};

			if (tPos.TableName != "PosArt") tPos.TableName = "PosArt";
			safeAdd(tPos);
			safeAdd(tTof);
			safeAdd(tRep);
			safeAdd(tIva);
			safeAdd(tLne);

			s = "SELECT * FROM GesOffArticoli INNER JOIN GesOffTestate ON GesOffArticoli.ofa_yea = GesOffTestate.oft_yea AND GesOffArticoli.ofa_cod = GesOffTestate.oft_cod WHERE GesOffArticoli.ofa_yea='" + strYea + "' AND GesOffArticoli.ofa_cod='" + strCod + "' AND GesOffArticoli.ofa_ann=0";
			DataTable t = _clsFun.FillTabSql("GesOffArticoli", s, bPrimo: false, _strConSql);
			if (t != null)
			{
				foreach (DataRow y in t.Rows)
				{
					if (y["ofa_art"] == DBNull.Value) continue;
					string sArt = y["ofa_art"].ToString().Trim();
					if (string.IsNullOrEmpty(sArt)) continue;

					DataRow[] j = tPos.Select("pos_art='" + sArt + "'");
					if (j.Length == 0)
					{
						DataRow x = tVar.NewRow();
						x["var_tip"] = _clsDef.VARPOS;
						x["var_num"] = _clsDef.COD03Z;
						x["var_art"] = sArt;
						x["var_dti"] = y["oft_dti"];
						x["var_dtv"] = y["oft_dti"];
						if (strSta == _clsDef.STADAC)
						{
							x["var_ori"] = "CHIUSURA OFFERTA";
						}
						FillPos(dSet, x);
						j = tPos.Select("pos_art='" + sArt + "'");
					}
					if (strSta != _clsDef.STADAC)
					{
						for (int i = 0; i < j.Length; i++)
						{
							FillOffArt(tTof, j[i], y);
						}
					}

					// Sincronizza anche le etichette se non già presenti
					if (tEti != null && tEti.Select("eti_art='" + sArt + "'").Length == 0)
					{
						DataRow xEti = tVar.NewRow();
						xEti["var_tip"] = _clsDef.VARETI;
						xEti["var_num"] = _clsDef.COD03Z;
						xEti["var_art"] = sArt;
						xEti["var_dti"] = y["oft_dti"];
						xEti["var_dtv"] = y["oft_dtf"];
						xEti["var_off"] = (strSta != _clsDef.STADAC) ? strCod : "";
						xEti["var_ori"] = (strSta == _clsDef.STADAC) ? "CHIUSURA OFFERTA" : ("Offerta " + strCod);
						FillEti(dSet, tEti, xEti, null);
					}
				}
			}
		}
		catch (Exception ex)
		{
			_clsFun.ErrorLog("frmGesVariazioni.FillVarOffArt", ex.Message);
		}
		finally
		{
			if (tPos != null && dSet.Tables.Contains(tPos.TableName))
			{
				dSet.Tables.Remove(tPos);
			}
			dSet.Tables.Clear();
			dSet.Dispose();
		}
	}

	private void FillOffArt(DataTable tabTof, DataRow rowPos, DataRow rowOfa)
	{
		rowPos["pos_off"] = rowOfa["oft_cod"];
		rowPos["pos_ori"] = rowOfa["oft_cod"];
		rowPos["pos_ost"] = rowOfa["oft_sta"];
		rowPos["pos_dti"] = rowOfa["oft_dti"];
		rowPos["pos_dtf"] = rowOfa["oft_dtf"];
		rowPos["pos_cam"] = rowOfa["oft_cam"];
		rowPos["pos_neg"] = rowOfa["oft_neg"];
		rowPos["pos_pun"] = rowOfa["ofa_pun"];
		rowPos["pos_tva"] = rowOfa["ofa_tip"];
		if ((string)rowPos["pos_tva"] == _clsDef.OFAPRZ || (string)rowPos["pos_tva"] == _clsDef.OFAMXN)
		{
			rowPos["pos_prv"] = rowOfa["ofa_val"];
		}
		rowPos["pos_xem"] = rowOfa["ofa_xem"];
		rowPos["pos_xen"] = rowOfa["ofa_xen"];
		rowPos["pos_mix"] = rowOfa["ofa_mix"];
		DataRow[] j2 = tabTof.Select("tab_cod='" + rowOfa["ofa_tip"]?.ToString() + "'");
		if (j2.Length != 0)
		{
			rowPos["pos_tvd"] = j2[0]["tab_des"];
			if ((string)rowOfa["ofa_tip"] == _clsDef.OFAMXN)
			{
				rowPos["pos_ori"] = "Off " + rowOfa["oft_cod"]?.ToString() + " (" + Convert.ToString(rowOfa["ofa_xem"]) + "x" + Convert.ToString(rowOfa["ofa_xen"]) + ")";
			}
			if ((string)rowPos["pos_tva"] == _clsDef.OFASCO)
			{
				rowPos["pos_ofd"] = Convert.ToInt16(rowOfa["ofa_val"]) + "%";
				rowPos["pos_ova"] = (decimal)rowOfa["ofa_val"];
			}
			else if ((string)rowPos["pos_tva"] == _clsDef.OFAMXN)
			{
				rowPos["pos_ofd"] = (decimal)rowOfa["ofa_xem"] + "x" + (decimal)rowOfa["ofa_xen"];
			}
		}
	}

	private void FillVarEti()
	{
		string s = "SELECT ";
		s += "GesVarEtichette.eti_nva, ";
		s += "GesVarEtichette.eti_art, ";
		s += "GesVarEtichette.eti_day, ";
		s += "GesVarEtichette.eti_ard, ";
		s += "GesVarEtichette.eti_prv, ";
		s += "GesVarEtichette.eti_ori, ";
		s += "GesVarEtichette.eti_sta, ";
		s += "GesVarEtichette.eti_tva, ";
		s += "GesVarEtichette.eti_inv, ";
		s += "GesVarEtichette.eti_msg, ";
		s += "TabStatoDivulgazioni.tab_des AS DesDiv ";
		s += "FROM GesVarEtichette ";
		s += "LEFT OUTER JOIN TabStatoDivulgazioni ON GesVarEtichette.eti_inv = TabStatoDivulgazioni.tab_cod ";
		s += "ORDER BY GesVarEtichette.eti_ard";
		DataTable t = _clsFun.FillTabSql("GesVarEtichette", s, bPrimo: false, _strConSql);
		dgv3.DataSource = t;
	}

	private void invioVariazioniACasseEBilanceToolStripMenuItem_Click(object sender, EventArgs e)
	{
		try
		{
			dgv1.CommitEdit(DataGridViewDataErrorContexts.Commit);
			CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager; if (cm != null && cm.Position >= 0 && dgv1.CurrentRow != null && cm.List[dgv1.CurrentRow.Index] is DataRowView)
			{
				cm.EndCurrentEdit();
			}
			DataTable tPos = (DataTable)dgv1.DataSource;
			bool hasADV = false;
			if (tPos != null)
			{
				foreach (DataRow r in tPos.Rows)
				{
					if (r.RowState != DataRowState.Deleted)
					{
						decimal oldP = ((r["pos_pvo"] != DBNull.Value) ? Convert.ToDecimal(r["pos_pvo"]) : 0m);
						decimal newP = ((r["pos_prv"] != DBNull.Value) ? Convert.ToDecimal(r["pos_prv"]) : 0m);
						string inv = ((r["pos_inv"] != DBNull.Value) ? r["pos_inv"].ToString() : "");
						if (oldP > 0m && newP != oldP && inv == _clsDef.DIVDAD)
						{
							hasADV = true;
							break;
						}
					}
				}
			}
			if (hasADV && MessageBox.Show("Sono presenti variazioni di prezzo provenienti da Terminalino.\r\n\r\nVuoi aggiornare automaticamente l'Anagrafica Articoli (listini di vendita) e accodare la stampa delle Etichette con i nuovi prezzi?", "AGGIORNAMENTO ANAGRAFICHE", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
			{
				frmWait.ShowWait("Aggiornamento anagrafica articoli e listini in corso... ATTENDERE.");
				UpdateFromADV(tPos);
				frmWait.CloseWait();
			}
		}
		catch (Exception ex)
		{
			frmWait.CloseWait();
			MessageBox.Show(this, "Errore durante il controllo preliminare: " + ex.Message, "ERRORE", MessageBoxButtons.OK, MessageBoxIcon.Hand);
			return;
		}
		menuStrip1.Enabled = false;
		frmWait.ShowWait("Generazione ed invio variazioni a Casse e Bilance in corso... ATTENDERE.");
		try
		{
			GenPos();
			Salva();
			try
			{
				string s = "DELETE FROM TmpDiv WHERE tmp_tip='ETI' AND tmp_sta='S'";
				_clsFun.SqlWrite(s, _strConSql);
			}
			catch (Exception ex)
			{
				_clsFun.ErrorLog("frmGesVariazioni.invioVariazioniACasseEBilance_TmpDivCleanup", ex.Message);
			}
			frmWait.CloseWait();
			MessageBox.Show(this, "Invio variazioni a Casse e Bilance completato con successo!", "INVIO CASSE E BILANCE", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
		}
		catch (Exception ex2)
		{
			frmWait.CloseWait();
			MessageBox.Show(this, "Errore durante l'invio variazioni: " + ex2.Message, "ERRORE", MessageBoxButtons.OK, MessageBoxIcon.Hand);
			_clsFun.ErrorLog("frmGesVariazioni.invioVariazioniACasseEBilanceToolStripMenuItem_Click", ex2.Message);
		}
		finally
		{
			frmWait.CloseWait();
			menuStrip1.Enabled = true;
		}
	}

	private void UpdateFromADV(DataTable tPos)
	{
		DataTable tEti = (DataTable)dgv3.DataSource;
		var priceUpdates = new Dictionary<string, decimal>();

		foreach (DataRow r in tPos.Rows)
		{
			if (r.RowState == DataRowState.Deleted)
			{
				continue;
			}
			decimal oldP = ((r["pos_pvo"] != DBNull.Value) ? Convert.ToDecimal(r["pos_pvo"]) : 0m);
			decimal newP = ((r["pos_prv"] != DBNull.Value) ? Convert.ToDecimal(r["pos_prv"]) : 0m);
			string inv = ((r["pos_inv"] != DBNull.Value) ? r["pos_inv"].ToString() : "");
			string artCod = ((r["pos_art"] != DBNull.Value) ? r["pos_art"].ToString() : "");
			if (!(oldP > 0m) || !(newP != oldP) || !(inv == _clsDef.DIVDAD) || string.IsNullOrEmpty(artCod))
			{
				continue;
			}
			priceUpdates[artCod] = newP;
			if (tEti != null)
			{
				DataRow[] existEti = tEti.Select("eti_art = '" + artCod.Replace("'", "''") + "' AND eti_inv = '" + _clsDef.DIVDAD + "'");
				if (existEti.Length == 0)
				{
					DataRow rEti = tEti.NewRow();
					rEti["eti_art"] = artCod;
					rEti["eti_day"] = DateTime.Today;
					rEti["eti_ard"] = ((r["pos_ard"] != DBNull.Value) ? r["pos_ard"].ToString() : "");
					rEti["eti_prv"] = newP;
					rEti["eti_ori"] = "ADV";
					rEti["eti_sta"] = "A";
					rEti["eti_inv"] = _clsDef.DIVDAD;
					tEti.Rows.Add(rEti);
				}
			}
		}

		if (priceUpdates.Count > 0)
		{
			_clsQry.AggiornaPrezzoVenditaConStoricoBatch(priceUpdates, _clsDef.LISPOS, _strConSql);
		}

		SalvaVar(_clsDef.VARETI);
	}

	private void GenPos()
	{
		string s = "SELECT * FROM AnaClienti WHERE cli_div='" + _clsDef.DIVDAD + "'";
		DataTable tCli = _clsFun.FillTabSql("AnaClienti", s, bPrimo: false, _strConSql);
		DataTable t = (DataTable)dgv1.DataSource;
		frmGesVarPos f = new frmGesVarPos();
		f._tabCli = tCli;
		f._tabPos = t;
		if (invioVariazioniACasseEBilanceToolStripMenuItem.Text == "Invio a terminalino")
		{
			f._strTip = "SOLOTERM";
		}
		f.ShowDialog(this);
		DataTable tOff = (DataTable)dgv2.DataSource;
		s = "SELECT tab_cod, tab_des, tab_art, tab_off FROM TabStato";
		DataTable tSta = _clsFun.FillTabSql("TabStato", s, bPrimo: false, _strConSql);
		if (tOff.Rows.Count > 0)
		{
			foreach (DataRow y in tOff.Rows)
			{
				s = (string)y["oft_cod"];
				if (((string)y["oft_sta"]).Trim() == _clsDef.STADAA)
				{
					y["oft_sta"] = _clsDef.STAATT;
				}
				else if (((string)y["oft_sta"]).Trim() == _clsDef.STADAC)
				{
					f.ChiudiOfferta(y);
					y["oft_sta"] = _clsDef.STACLO;
				}
				DataRow[] j = tSta.Select("tab_cod='" + y["oft_sta"]?.ToString() + "'");
				if (j.Length != 0)
				{
					y["OftStd"] = (string)j[0]["tab_des"];
				}
			}
		}
		s = "UPDATE AnaClienti SET cli_div='' WHERE cli_div='" + _clsDef.DIVDAD + "'";
		_clsFun.SqlWrite(s, _strConSql);
	}

	private void Salva()
	{
		SalvaVar(_clsDef.VARPOS);
		SalvaVar(_clsDef.VARETI);
		SalvaOff();
	}

	private void SalvaVar(string strTip)
	{
		DataTable t = (DataTable)dgv1.DataSource;
		if (strTip == _clsDef.VARETI)
		{
			t = (DataTable)dgv3.DataSource;
		}
		if (t == null || t.Rows.Count <= 0)
		{
			return;
		}
		int iNum = 0;
		string s = "SELECT * FROM GesVariazioni WHERE var_tip='" + strTip + "' AND var_num <> '000' ORDER BY var_dtv DESC, var_num DESC";
		DataTable tVarExist = _clsFun.FillTabSql("GesVariazioni", s, bPrimo: true, _strConSql);
		if (tVarExist.Rows.Count > 0)
		{
			iNum = Convert.ToInt32(tVarExist.Rows[0]["var_num"]);
		}
		iNum++;
		if (iNum > 999)
		{
			iNum = 1;
		}
		DateTime dd = DateTime.Today.AddDays(-90.0);
		s = "DELETE FROM GesVariazioni WHERE var_dtv < " + _clsFun.DaySql(dd);
		_clsFun.SqlWrite(s, _strConSql);
		s = "DELETE FROM GesVariazioni WHERE var_tip='" + strTip + "' AND var_num='" + iNum.ToString("000") + "'";
		_clsFun.SqlWrite(s, _strConSql);
		string sFldInv = "pos_inv";
		string sFldArt = "pos_art";
		if (strTip == _clsDef.VARETI)
		{
			sFldInv = "eti_inv";
			sFldArt = "eti_art";
		}
		s = "SELECT * FROM GesVariazioni WHERE var_num='000'";
		DataTable tVar = _clsFun.FillTabSql("GesVariazioni", s, bPrimo: false, _strConSql);
		HashSet<string> existingVarArts = new HashSet<string>();
		foreach (DataRow rv in tVar.Rows)
		{
			if (rv["var_tip"] != DBNull.Value && rv["var_art"] != DBNull.Value)
			{
				string tip = rv["var_tip"].ToString().Trim();
				string art = rv["var_art"].ToString().Trim().PadLeft(7, '0');
				existingVarArts.Add(tip + "|" + art);
			}
		}
		if (progressBar1 != null && t.Rows.Count > 0)
		{
			progressBar1.Minimum = 0;
			progressBar1.Maximum = t.Rows.Count;
			progressBar1.Value = 0;
		}
		StringBuilder sbBatch = new StringBuilder();
		int batchCount = 0;
		int rowIdx = 0;
		foreach (DataRow y in t.Rows)
		{
			rowIdx++;
			if (progressBar1 != null && rowIdx % 100 == 0 && rowIdx <= progressBar1.Maximum)
			{
				progressBar1.Value = rowIdx;
				Application.DoEvents();
			}
			string artCode = (y[sFldArt] != DBNull.Value) ? y[sFldArt].ToString().Trim().PadLeft(7, '0') : "";
			if (string.IsNullOrEmpty(artCode)) continue;

			if ((string)y[sFldInv] != _clsDef.DIVDAD)
			{
				s = "UPDATE GesVariazioni SET ";
				if ((string)y[sFldInv] != _clsDef.DIVSOS)
				{
					s = s + "var_num='" + iNum.ToString("000") + "', ";
				}
				s = ((!(strTip == _clsDef.VARPOS) || !(((string)y["pos_msg"]).Trim() != "")) ? (s + "var_inv='" + y[sFldInv]?.ToString() + "', ") : (s + "var_inv='" + _clsDef.DIVDAD + "', "));
				s = s + "var_dtv=" + _clsFun.DaySql(DateTime.Today) + " ";
				s = s + "WHERE var_tip='" + strTip + "' AND var_num='000' AND var_art='" + artCode + "'";
				sbBatch.AppendLine(s + ";");
				batchCount++;
			}
			else
			{
				DateTime dDti = DateTime.Today;
				DateTime dDtv = DateTime.Today;
				string sOri = "";
				string sOff = "";
				decimal dEtq = default(decimal);
				string sEan = "";
				if (strTip == "POS")
				{
					sOff = ((y.Table.Columns.Contains("pos_off") && y["pos_off"] != DBNull.Value) ? ((string)y["pos_off"]) : "");
					sOri = ((y.Table.Columns.Contains("pos_ori") && y["pos_ori"] != DBNull.Value) ? ((string)y["pos_ori"]) : "");
				}
				else if (strTip == "ETI")
				{
					sOff = ((y.Table.Columns.Contains("eti_off") && y["eti_off"] != DBNull.Value) ? ((string)y["eti_off"]) : "");
					sOri = ((y.Table.Columns.Contains("eti_ori") && y["eti_ori"] != DBNull.Value) ? ((string)y["eti_ori"]) : "");
					sEan = ((y.Table.Columns.Contains("eti_ean") && y["eti_ean"] != DBNull.Value) ? ((string)y["eti_ean"]) : "");
				}
				string varKey = strTip + "|" + artCode;
				if (!existingVarArts.Contains(varKey))
				{
					existingVarArts.Add(varKey);
					DataRow x = tVar.NewRow();
					x["var_tip"] = strTip;
					x["var_num"] = "000";
					x["var_art"] = artCode;
					x["var_inv"] = _clsDef.DIVDAD;
					x["var_dti"] = dDti;
					x["var_dtv"] = dDtv;
					x["var_ori"] = sOri;
					x["var_off"] = sOff;
					x["var_etq"] = dEtq;
					x["var_ean"] = sEan;
					s = _clsFun.SqlInsertRow("GesVariazioni", tVar, x);
					sbBatch.AppendLine(s + ";");
					batchCount++;
				}
			}
			if (batchCount >= 100)
			{
				_clsFun.SqlWrite(sbBatch.ToString(), _strConSql);
				sbBatch.Clear();
				batchCount = 0;
			}
		}
		if (sbBatch.Length > 0)
		{
			_clsFun.SqlWrite(sbBatch.ToString(), _strConSql);
			sbBatch.Clear();
		}
	}

	private void SalvaOff()
	{
		string s = "";
		s = "SELECT ";
		s += "GesOffTestate.* ";
		s += "FROM GesOffTestate ";
		s += "WHERE ";
		s = s + "GesOffTestate.oft_sta <> '" + _clsDef.STACLO + "' ";
		s += "ORDER BY oft_dti DESC";
		DataTable t = _clsFun.FillTabSql("GesOffTestate", s, bPrimo: false, _strConSql);
		DataTable tOff = (DataTable)dgv2.DataSource;
		if (tOff == null || tOff.Rows.Count <= 0)
		{
			return;
		}
		foreach (DataRow y in tOff.Rows)
		{
			DataRow[] j = t.Select("oft_yea='" + y["oft_yea"]?.ToString() + "' AND oft_cod='" + y["oft_cod"]?.ToString() + "'");
			if (j.Length != 0 && (string)j[0]["oft_sta"] != (string)y["oft_sta"])
			{
				s = "UPDATE GesOffTestate SET oft_sta='" + (string)y["oft_sta"] + "' WHERE ";
				s = s + "oft_yea='" + y["oft_yea"]?.ToString() + "' AND oft_cod='" + y["oft_cod"]?.ToString() + "'";
				_clsFun.SqlWrite(s, _strConSql);
			}
		}
	}

	private void dgv_CurrentCellDirtyStateChanged(object sender, EventArgs e)
	{
	}

	private void VarEtichette(string strVar)
	{
		foreach (DataRow y in ((DataTable)dgv3.DataSource).Rows)
		{
			y["eti_inv"] = strVar;
		}
	}

	private void VarPos(string strVar)
	{
		foreach (DataRow y in ((DataTable)dgv1.DataSource).Rows)
		{
			y["pos_inv"] = strVar;
		}
	}

	private void Recupera(string strTip)
	{
		frmGesVarStorico f = new frmGesVarStorico();
		f._strVarTip = strTip;
		f.ShowDialog(this);
		if (f._strVarNum != "")
		{
			string[] a = f._strVarNum.Split(';');
			string[] array = a;
			foreach (string sNum in array)
			{
				FillVarPos(strTip, sNum);
			}
		}
	}

	private void annullaTutteLeEtichetteToolStripMenuItem_Click(object sender, EventArgs e)
	{
		VarEtichette(_clsDef.DIVANN);
	}

	private void attivaTutteLeEtichetteToolStripMenuItem1_Click(object sender, EventArgs e)
	{
		VarEtichette(_clsDef.DIVDAD);
	}

	private void recuperoEtichetteToolStripMenuItem1_Click(object sender, EventArgs e)
	{
		Recupera(_clsDef.VARETI);
	}

	private void annullaTutteLeVariazioniToolStripMenuItem_Click(object sender, EventArgs e)
	{
		VarPos(_clsDef.DIVANN);
	}

	private void attivaTutteLeVariazioniToolStripMenuItem_Click(object sender, EventArgs e)
	{
		VarPos(_clsDef.DIVDAD);
	}

	private void recuperoVariazioniToolStripMenuItem_Click(object sender, EventArgs e)
	{
		Recupera(_clsDef.VARPOS);
	}

	private void svuotaElencoVariazioniToolStripMenuItem_Click(object sender, EventArgs e)
	{
		DialogResult res1 = MessageBox.Show("Sei sicuro di voler svuotare l'intero elenco variazioni (etichette, casse e bilance)?", "CONFERMA SVUOTAMENTO VARIAZIONI", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button2);
		if (res1 != DialogResult.Yes)
		{
			return;
		}
		DialogResult res2 = MessageBox.Show("ATTENZIONE: L'operazione cancellerà DEFINITIVAMENTE tutte le variazioni etichette e tutte le variazioni per casse e bilance senza alcuna possibilità di recupero!\n\nSei VERAMENTE sicuro di voler procedere con l'eliminazione totale?", "CONFERMA DEFINITIVA - NESSUN RECUPERO POSSIBILE", MessageBoxButtons.YesNo, MessageBoxIcon.Hand, MessageBoxDefaultButton.Button2);
		if (res2 != DialogResult.Yes)
		{
			return;
		}
		try
		{
			string s = "DELETE FROM GesVariazioni";
			_clsFun.SqlWrite(s, _strConSql);
			ClearArticleCache();
			FillVarPos("", "");
			FillOfferte();
			MessageBox.Show("L'elenco delle variazioni etichette, casse e bilance è stato svuotato definitivamente.", "SVUOTAMENTO COMPLETATO", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
		}
		catch (Exception ex)
		{
			_clsFun.ErrorLog("frmGesVariazioni.svuotaElencoVariazioniToolStripMenuItem_Click", ex.Message);
			MessageBox.Show("Errore durante lo svuotamento delle variazioni: " + ex.Message, "ERRORE", MessageBoxButtons.OK, MessageBoxIcon.Hand);
		}
	}

	private void terminalinoToolStripMenuItem_Click(object sender, EventArgs e)
	{
		frmGesVarTerm f = new frmGesVarTerm();
		f.ShowDialog(this);
		if (!f._bolOk)
		{
			return;
		}
		DataTable tPos = new clsGenTabTmp().TabTmpPos("PosArt");
		string s = "SELECT * FROM GesVariazioni WHERE var_num='" + _clsDef.COD03Z + "'";
		DataTable tVar = _clsFun.FillTabSql("GesVariazioni", s, bPrimo: false, _strConSql);
		s = "SELECT lia_art FROM GesLisAcquisto WHERE lia_for='" + f._strDiv + "'";
		DataTable tLia = _clsFun.FillTabSql("GesLisAcquisto", s, bPrimo: false, _strConSql);
		s = "SELECT * FROM TabReparti WHERE tab_ann=0";
		DataTable tRep = _clsFun.FillTabSql("TabReparti", s, bPrimo: false, _strConSql);
		s = "SELECT * FROM TabIva WHERE tab_ann=0";
		DataTable tIva = _clsFun.FillTabSql("TabIva", s, bPrimo: false, _strConSql);
		DataTable tLne = _clsFun.FillTabSql("TabLisNegozi", "SELECT tab_lis FROM TabNegozi WHERE tab_tip = 'L'", bPrimo: false, _strConSql);
		string sLisNeg = ((tLne != null && tLne.Rows.Count > 0) ? ((string)tLne.Rows[0]["tab_lis"]) : "");
		DataSet dSet = new DataSet();
		dSet.Tables.Add(tPos);
		dSet.Tables.Add(tRep);
		dSet.Tables.Add(tIva);
		LookupCache cache = new LookupCache();
		foreach (DataRow r in tRep.Rows)
		{
			cache.Reparti[r["tab_cod"].ToString()] = r;
		}
		foreach (DataRow r2 in tIva.Rows)
		{
			cache.Iva[r2["tab_cod"].ToString()] = r2;
		}
		HashSet<string> addedPos = new HashSet<string>();
		DataTable tPreloadVar = new DataTable();
		tPreloadVar.Columns.Add("var_art", typeof(string));
		foreach (DataRow r3 in tLia.Rows)
		{
			DataRow nr = tPreloadVar.NewRow();
			nr["var_art"] = r3["lia_art"];
			tPreloadVar.Rows.Add(nr);
		}
		PreloadArticlesCache(tPreloadVar, sLisNeg);
		int totalLia = tLia.Rows.Count;
		progressBar1.Value = 0;
		progressBar1.Maximum = ((totalLia <= 0) ? 1 : totalLia);
		progressBar1.Minimum = 0;
		int iCount = 0;
		foreach (DataRow y in tLia.Rows)
		{
			iCount++;
			if (iCount % 50 == 0 || iCount == totalLia)
			{
				progressBar1.Value = Math.Min(iCount, progressBar1.Maximum);
				Application.DoEvents();
			}
			DataRow x = tVar.NewRow();
			x["var_tip"] = _clsDef.VARPOS;
			x["var_num"] = _clsDef.COD03Z;
			x["var_art"] = y["lia_art"];
			x["var_ori"] = "Impianto terminalino";
			x["var_dti"] = DateTime.Today;
			x["var_dtv"] = DateTime.Today;
			FillPos(dSet, x, addedPos, cache);
		}
		ClearArticleCache();
		dgv1.DataSource = tPos;
		lblVarCnt.Text = tPos.Rows.Count.ToString();
		dSet.Tables.Clear();
		dSet.Clear();
		dSet.Dispose();
		invioVariazioniACasseEBilanceToolStripMenuItem.Text = "Invio a terminalino";
	}

	private void impiantoCassaToolStripMenuItem_Click(object sender, EventArgs e)
	{
		frmGesVarPosImpianto f = new frmGesVarPosImpianto();
		f.ShowDialog();
		if (!f._bolOk)
		{
			return;
		}
		frmWait.ShowWait("GENERAZIONE IMPIANTO CASSA IN CORSO... ATTENDERE.");
		try
		{
			ClearArticleCache();
			DataTable tPos = new clsGenTabTmp().TabTmpPos("PosArt");
			string s = "SELECT * FROM GesVariazioni WHERE var_num='" + _clsDef.COD03Z + "'";
			DataTable tVar = _clsFun.FillTabSql("GesVariazioni", s, bPrimo: false, _strConSql);
			s = "SELECT art_cod FROM AnaArticoli ";
			if (!f._bolAll)
			{
				s = s + "WHERE art_sta='" + _clsDef.STAATT + "'";
			}
			DataTable tArt = _clsFun.FillTabSql("AnaArticoli", s, bPrimo: false, _strConSql);
			s = "SELECT * FROM TabReparti WHERE tab_ann=0";
			DataTable tRep = _clsFun.FillTabSql("TabReparti", s, bPrimo: false, _strConSql);
			s = "SELECT * FROM TabIva WHERE tab_ann=0";
			DataTable tIva = _clsFun.FillTabSql("TabIva", s, bPrimo: false, _strConSql);
			DataTable tLne = _clsFun.FillTabSql("TabLisNegozi", "SELECT tab_lis FROM TabNegozi WHERE tab_tip = 'L'", bPrimo: false, _strConSql);
			string sLisNeg = ((tLne != null && tLne.Rows.Count > 0) ? ((string)tLne.Rows[0]["tab_lis"]) : "");
			DataSet dSet = new DataSet();
			dSet.Tables.Add(tPos);
			dSet.Tables.Add(tRep);
			dSet.Tables.Add(tIva);
			LookupCache cache = new LookupCache();
			foreach (DataRow r in tRep.Rows)
			{
				cache.Reparti[r["tab_cod"].ToString()] = r;
			}
			foreach (DataRow r2 in tIva.Rows)
			{
				cache.Iva[r2["tab_cod"].ToString()] = r2;
			}
			HashSet<string> addedPos = new HashSet<string>();
			DataTable tPreloadVar = new DataTable();
			tPreloadVar.Columns.Add("var_art", typeof(string));
			foreach (DataRow r3 in tArt.Rows)
			{
				DataRow nr = tPreloadVar.NewRow();
				nr["var_art"] = r3["art_cod"];
				tPreloadVar.Rows.Add(nr);
			}
			PreloadArticlesCache(tPreloadVar, sLisNeg);
			int totalArt = tArt.Rows.Count;
			progressBar1.Value = 0;
			progressBar1.Maximum = ((totalArt <= 0) ? 1 : totalArt);
			progressBar1.Minimum = 0;
			int iCount = 0;
			string origTitle = Text;
			foreach (DataRow y in tArt.Rows)
			{
				iCount++;
				if (iCount % 50 == 0 || iCount == totalArt)
				{
					int pct = (int)((double)iCount / (double)totalArt * 100.0);
					progressBar1.Value = Math.Min(iCount, progressBar1.Maximum);
					string msg = $"Impianto Cassa: {pct}% ({iCount}/{totalArt})";
					lblVarCnt.Text = msg;
					Text = "Gestione Variazioni - " + msg;
					frmWait.UpdateText($"GENERAZIONE IMPIANTO CASSA: {pct}% ({iCount}/{totalArt})");
					Application.DoEvents();
				}
				DataRow x = tVar.NewRow();
				x["var_tip"] = _clsDef.VARPOS;
				x["var_num"] = _clsDef.COD03Z;
				x["var_art"] = y["art_cod"];
				x["var_ori"] = "Impianto pos";
				x["var_dti"] = DateTime.Today;
				x["var_dtv"] = DateTime.Today;
				FillPos(dSet, x, addedPos, cache);
			}
			ClearArticleCache();
			dgv1.DataSource = tPos;
			lblVarCnt.Text = tPos.Rows.Count.ToString();
			Text = origTitle;
			dSet.Tables.Clear();
			dSet.Clear();
			dSet.Dispose();
		}
		finally
		{
			frmWait.CloseWait();
		}
	}

	private void impiantoEtichetteToolStripMenuItem_Click(object sender, EventArgs e)
	{
		frmGesVarPosImpianto f = new frmGesVarPosImpianto();
		f.Text = "Impianto Etichette";
		f.ShowDialog();
		if (!f._bolOk)
		{
			return;
		}
		frmWait.ShowWait("GENERAZIONE IMPIANTO ETICHETTE IN CORSO... ATTENDERE.");
		try
		{
			ClearArticleCache();
			DataTable tEti = new clsGenTabTmp().TabTmpEti("EtiArt");
			string s = "SELECT * FROM GesVariazioni WHERE var_num='" + _clsDef.COD03Z + "'";
			DataTable tVar = _clsFun.FillTabSql("GesVariazioni", s, bPrimo: false, _strConSql);
			s = "SELECT art_cod FROM AnaArticoli ";
			if (!f._bolAll)
			{
				s = s + "WHERE art_sta='" + _clsDef.STAATT + "'";
			}
			DataTable tArt = _clsFun.FillTabSql("AnaArticoli", s, bPrimo: false, _strConSql);
			s = "SELECT * FROM TabReparti WHERE tab_ann=0";
			DataTable tRep = _clsFun.FillTabSql("TabReparti", s, bPrimo: false, _strConSql);
			s = "SELECT * FROM TabIva WHERE tab_ann=0";
			DataTable tIva = _clsFun.FillTabSql("TabIva", s, bPrimo: false, _strConSql);
			DataTable tOffEti = _clsQry.OfsSeekEti(DateTime.Today.AddDays(-15.0));
			DataTable tLne = _clsFun.FillTabSql("TabLisNegozi", "SELECT tab_lis FROM TabNegozi WHERE tab_tip = 'L'", bPrimo: false, _strConSql);
			string sLisNeg = ((tLne != null && tLne.Rows.Count > 0) ? ((string)tLne.Rows[0]["tab_lis"]) : "");
			DataSet dSet = new DataSet();
			dSet.Tables.Add(tEti);
			dSet.Tables.Add(tRep);
			dSet.Tables.Add(tIva);
			LookupCache cache = new LookupCache();
			foreach (DataRow r in tRep.Rows)
			{
				cache.Reparti[r["tab_cod"].ToString()] = r;
			}
			foreach (DataRow r2 in tIva.Rows)
			{
				cache.Iva[r2["tab_cod"].ToString()] = r2;
			}
			foreach (DataRow r3 in tOffEti.Rows)
			{
				cache.OffTes[r3["oft_cod"].ToString() + "|" + r3["ofa_art"].ToString()] = r3;
			}
			HashSet<string> addedEti = new HashSet<string>();
			DataTable tPreloadVar = new DataTable();
			tPreloadVar.Columns.Add("var_art", typeof(string));
			foreach (DataRow r4 in tArt.Rows)
			{
				DataRow nr = tPreloadVar.NewRow();
				nr["var_art"] = r4["art_cod"];
				tPreloadVar.Rows.Add(nr);
			}
			PreloadArticlesCache(tPreloadVar, sLisNeg);
			int totalArt = tArt.Rows.Count;
			progressBar1.Value = 0;
			progressBar1.Maximum = ((totalArt <= 0) ? 1 : totalArt);
			progressBar1.Minimum = 0;
			int iCount = 0;
			string origTitle = Text;
			foreach (DataRow y in tArt.Rows)
			{
				iCount++;
				if (iCount % 50 == 0 || iCount == totalArt)
				{
					int pct = (int)((double)iCount / (double)totalArt * 100.0);
					progressBar1.Value = Math.Min(iCount, progressBar1.Maximum);
					string msg = $"Impianto Etichette: {pct}% ({iCount}/{totalArt})";
					lblEtiCnt.Text = msg;
					Text = "Gestione Variazioni - " + msg;
					frmWait.UpdateText($"GENERAZIONE IMPIANTO ETICHETTE: {pct}% ({iCount}/{totalArt})");
					Application.DoEvents();
				}
				DataRow x = tVar.NewRow();
				x["var_tip"] = _clsDef.VARETI;
				x["var_num"] = _clsDef.COD03Z;
				x["var_art"] = y["art_cod"];
				x["var_ori"] = "Impianto etichette";
				x["var_dti"] = DateTime.Today;
				x["var_dtv"] = DateTime.Today;
				FillEti(dSet, tEti, x, tOffEti, addedEti, cache);
			}
			ClearArticleCache();
			dgv3.DataSource = tEti;
			lblEtiCnt.Text = tEti.Rows.Count.ToString();
			Text = origTitle;
			dSet.Tables.Clear();
			dSet.Clear();
			dSet.Dispose();
		}
		finally
		{
			frmWait.CloseWait();
		}
	}

	private void bilanciaToolStripMenuItem_Click(object sender, EventArgs e)
	{
		frmGesVarBilImpianto f = new frmGesVarBilImpianto();
		f.ShowDialog();
		if (!f._bolOk)
		{
			return;
		}
		frmWait.ShowWait("GENERAZIONE IMPIANTO BILANCE IN CORSO... ATTENDERE.");
		try
		{
			ClearArticleCache();
			DataTable tPos = new clsGenTabTmp().TabTmpPos("PosArt");
			string s = "SELECT * FROM GesVariazioni WHERE var_num='" + _clsDef.COD03Z + "'";
			DataTable tVar = _clsFun.FillTabSql("GesVariazioni", s, bPrimo: false, _strConSql);
			s = "SELECT art_cod FROM AnaArticoli ";
			if (!f._bolAll)
			{
				s = s + "WHERE art_sta='" + _clsDef.STAATT + "' AND art_bil=1 ";
				if (f._strReb != "")
				{
					s = s + "AND art_reb='" + f._strReb + "'";
				}
			}
			DataTable tArt = _clsFun.FillTabSql("AnaArticoli", s, bPrimo: false, _strConSql);
			s = "SELECT * FROM TabReparti WHERE tab_ann=0";
			DataTable tRep = _clsFun.FillTabSql("TabReparti", s, bPrimo: false, _strConSql);
			s = "SELECT * FROM TabIva WHERE tab_ann=0";
			DataTable tIva = _clsFun.FillTabSql("TabIva", s, bPrimo: false, _strConSql);
			DataTable tLne = _clsFun.FillTabSql("TabLisNegozi", "SELECT tab_lis FROM TabNegozi WHERE tab_tip = 'L'", bPrimo: false, _strConSql);
			string sLisNeg = ((tLne != null && tLne.Rows.Count > 0) ? ((string)tLne.Rows[0]["tab_lis"]) : "");
			DataSet dSet = new DataSet();
			dSet.Tables.Add(tPos);
			dSet.Tables.Add(tRep);
			dSet.Tables.Add(tIva);
			LookupCache cache = new LookupCache();
			foreach (DataRow r in tRep.Rows)
			{
				cache.Reparti[r["tab_cod"].ToString()] = r;
			}
			foreach (DataRow r2 in tIva.Rows)
			{
				cache.Iva[r2["tab_cod"].ToString()] = r2;
			}
			HashSet<string> addedPos = new HashSet<string>();
			DataTable tPreloadVar = new DataTable();
			tPreloadVar.Columns.Add("var_art", typeof(string));
			foreach (DataRow r3 in tArt.Rows)
			{
				DataRow nr = tPreloadVar.NewRow();
				nr["var_art"] = r3["art_cod"];
				tPreloadVar.Rows.Add(nr);
			}
			PreloadArticlesCache(tPreloadVar, sLisNeg);
			int totalArt = tArt.Rows.Count;
			progressBar1.Value = 0;
			progressBar1.Maximum = ((totalArt <= 0) ? 1 : totalArt);
			progressBar1.Minimum = 0;
			int iCount = 0;
			string origTitle = Text;
			foreach (DataRow y in tArt.Rows)
			{
				iCount++;
				if (iCount % 50 == 0 || iCount == totalArt)
				{
					int pct = (int)((double)iCount / (double)totalArt * 100.0);
					progressBar1.Value = Math.Min(iCount, progressBar1.Maximum);
					string msg = $"Impianto Bilance: {pct}% ({iCount}/{totalArt})";
					lblVarCnt.Text = msg;
					Text = "Gestione Variazioni - " + msg;
					frmWait.UpdateText($"GENERAZIONE IMPIANTO BILANCE: {pct}% ({iCount}/{totalArt})");
					Application.DoEvents();
				}
				DataRow x = tVar.NewRow();
				x["var_tip"] = _clsDef.VARPOS;
				x["var_num"] = _clsDef.COD03Z;
				x["var_art"] = y["art_cod"];
				x["var_ori"] = "Impianto pos";
				x["var_dti"] = DateTime.Today;
				x["var_dtv"] = DateTime.Today;
				FillPos(dSet, x, addedPos, cache);
			}
			ClearArticleCache();
			dgv1.DataSource = tPos;
			lblVarCnt.Text = tPos.Rows.Count.ToString();
			Text = origTitle;
			dSet.Tables.Clear();
			dSet.Clear();
			dSet.Dispose();
		}
		finally
		{
			frmWait.CloseWait();
		}
	}

	private void importTerminalinoToolStripMenuItem_Click(object sender, EventArgs e)
	{
		ImpTerm("ETI");
	}

	private void importTerminalinoToolStripMenuItem1_Click(object sender, EventArgs e)
	{
		ImpTerm("ALL");
	}

	private void importTerminalinoADVToolStripMenuItem_Click(object sender, EventArgs e)
	{
		ImpTerm("ADV");
	}

	private void ImpTerm(string strTip)
	{
		string s = "";
		string sMsg = "";
		DataTable tTer = new clsGenTabTmp().TabTmpTerRil("TabTer");
		frmGesImpTerm f = new frmGesImpTerm();
		f._strTip = ((strTip == "ALL") ? "ETI" : strTip);
		f._tabTer = tTer;
		f.ShowDialog();
		if (f._tabImp == null)
		{
			return;
		}
		ClearArticleCache();
		DataTable tPos = (DataTable)dgv1.DataSource;
		DataTable tTgr = _clsFun.FillTabSql("TabTipoGrammatura", "SELECT * FROM TabTipoGrammatura", bPrimo: false, _strConSql);
		DataTable tCal = _clsFun.FillTabSql("TabCalibro", "SELECT * FROM TabCalibro", bPrimo: false, _strConSql);
		DataTable tCat = _clsFun.FillTabSql("TabCategoria", "SELECT * FROM TabCategoria", bPrimo: false, _strConSql);
		DataTable tOri = _clsFun.FillTabSql("TabOrigine", "SELECT * FROM TabOrigine", bPrimo: false, _strConSql);
		DataTable tPro = _clsQry.LivPromo();
		s = "SELECT * FROM TabReparti WHERE tab_ann=0";
		DataTable tRep = _clsFun.FillTabSql("TabReparti", s, bPrimo: false, _strConSql);
		s = "SELECT * FROM TabIva WHERE tab_ann=0";
		DataTable tIva = _clsFun.FillTabSql("TabIva", s, bPrimo: false, _strConSql);
		DataTable tLne = _clsFun.FillTabSql("TabLisNegozi", "SELECT tab_lis FROM TabNegozi WHERE tab_tip = 'L'", bPrimo: false, _strConSql);
		string sLisNeg = ((tLne != null && tLne.Rows.Count > 0) ? ((string)tLne.Rows[0]["tab_lis"]) : "");
		DataTable tOffEti = _clsQry.OfsSeekEti(DateTime.Today.AddDays(-15.0));
		DataSet dSet = new DataSet();
		dSet.Tables.Add(tPos.Clone());
		dSet.Tables.Add(tTgr);
		dSet.Tables.Add(tCal);
		dSet.Tables.Add(tCat);
		dSet.Tables.Add(tOri);
		dSet.Tables.Add(tPro);
		DataTable t = f._tabImp.Copy();
		DataTable tEti = (DataTable)dgv3.DataSource;
		s = "SELECT * FROM GesVariazioni WHERE var_num='" + _clsDef.COD03Z + "'";
		DataTable tVar = _clsFun.FillTabSql("GesVariazioni", s, bPrimo: false, _strConSql);
		DataTable tPreloadVar = new DataTable();
		tPreloadVar.Columns.Add("var_art", typeof(string));
		foreach (DataRow y in t.Rows)
		{
			if (((string)y["ord_art"]).Trim() == "")
			{
				DataTable tTmp = _clsQry.ArtSeek((string)y["ord_ean"], "");
				if (tTmp.Rows.Count > 0)
				{
					y["ord_art"] = ((string)tTmp.Rows[0]["tmp_art"]).PadLeft(7, Convert.ToChar("0"));
					if ((string)tTmp.Rows[0]["tmp_sta"] == _clsDef.STANOA)
					{
						_clsQry.ArtAttiva((string)tTmp.Rows[0]["tmp_art"]);
					}
				}
			}
			if (((string)y["ord_art"]).Trim() != "")
			{
				DataRow nr = tPreloadVar.NewRow();
				nr["var_art"] = ((string)y["ord_art"]).PadLeft(7, Convert.ToChar("0"));
				tPreloadVar.Rows.Add(nr);
			}
		}
		PreloadArticlesCache(tPreloadVar, sLisNeg);
		HashSet<string> addedPos = new HashSet<string>();
		HashSet<string> addedEti = new HashSet<string>();
		LookupCache cache = new LookupCache();
		foreach (DataRow r in tRep.Rows)
		{
			cache.Reparti[r["tab_cod"].ToString()] = r;
		}
		foreach (DataRow r2 in tIva.Rows)
		{
			cache.Iva[r2["tab_cod"].ToString()] = r2;
		}
		foreach (DataRow r3 in tTgr.Rows)
		{
			cache.Tgr[r3["tab_cod"].ToString()] = r3;
		}
		foreach (DataRow r4 in tCal.Rows)
		{
			cache.Cal[r4["tab_cod"].ToString()] = r4;
		}
		foreach (DataRow r5 in tCat.Rows)
		{
			cache.Cat[r5["tab_cod"].ToString()] = r5;
		}
		foreach (DataRow r6 in tOri.Rows)
		{
			cache.Ori[r6["tab_cod"].ToString()] = r6;
		}
		foreach (DataRow r7 in tOffEti.Rows)
		{
			cache.OffTes[r7["oft_cod"].ToString() + "|" + r7["ofa_art"].ToString()] = r7;
		}
		progressBar1.Value = 0;
		progressBar1.Maximum = t.Rows.Count;
		progressBar1.Minimum = 0;
		Dictionary<string, decimal> dictOldPrice = new Dictionary<string, decimal>();
		Dictionary<string, decimal> dictCost = new Dictionary<string, decimal>();
		if (strTip == "ADV")
		{
			List<string> codList = new List<string>();
			foreach (DataRow y2 in t.Rows)
			{
				string c = y2["ord_art"].ToString().Trim().PadLeft(7, '0');
				if (!string.IsNullOrEmpty(c))
				{
					codList.Add("'" + c.Replace("'", "''") + "'");
				}
			}
			if (codList.Count > 0)
			{
				string todaySql = _clsFun.DaySql(DateTime.Today);
				string inClause = string.Join(",", codList.Distinct());
				string sql = "SELECT art_cod, ISNULL(c.lia_cos, 0) AS true_cos, ISNULL(v.liv_prv, 0) AS true_prv \r\n                                       FROM AnaArticoli \r\n                                       OUTER APPLY (SELECT TOP 1 lia_cos FROM GesLisAcquisto WHERE lia_art = art_cod AND (lia_ann=0 OR lia_ann IS NULL) ORDER BY lia_dti DESC) c \r\n                                       OUTER APPLY (SELECT TOP 1 liv_prv FROM GesLisVendita WHERE liv_art = art_cod AND liv_lis = '" + _clsDef.LISPOS + "' AND (liv_ann=0 OR liv_ann IS NULL) ORDER BY liv_dti DESC) v \r\n                                       WHERE art_cod IN (" + inClause + ")";
				DataTable tReal = _clsFun.FillTabSql("RealPrices", sql, bPrimo: false, _strConSql);
				foreach (DataRow r8 in tReal.Rows)
				{
					dictCost[r8["art_cod"].ToString()] = Convert.ToDecimal(r8["true_cos"]);
					dictOldPrice[r8["art_cod"].ToString()] = Convert.ToDecimal(r8["true_prv"]);
				}
			}
		}
		foreach (DataRow y3 in t.Rows)
		{
			progressBar1.Increment(1);
			Application.DoEvents();
			if (((string)y3["ord_art"]).Trim() == "")
			{
				sMsg = sMsg + "Non trovato ean " + (string)y3["ord_ean"] + _clsDef.CRLF;
				continue;
			}
			DataRow x = tVar.NewRow();
			x["var_tip"] = _clsDef.VARETI;
			x["var_num"] = _clsDef.COD03Z;
			x["var_art"] = ((string)y3["ord_art"]).PadLeft(7, Convert.ToChar("0"));
			x["var_ori"] = "Da terminalino";
			x["var_dti"] = DateTime.Today;
			x["var_ean"] = ((string)y3["ord_ean"]).Trim();
			x["var_dtv"] = DateTime.Today;
			if (strTip == "ETI")
			{
				FillEti(dSet, tEti, x, tOffEti, addedEti, cache);
				continue;
			}
			FillPos(dSet, x, addedPos, cache);
			FillEti(dSet, tEti, x, tOffEti, addedEti, cache);
			if (!(strTip == "ADV"))
			{
				continue;
			}
			decimal newPrice = Convert.ToDecimal(y3["ord_qta"]);
			string currentArt = x["var_art"].ToString();
			if (dSet.Tables.Contains("PosArt"))
			{
				foreach (DataRow pRow in dSet.Tables["PosArt"].Rows)
				{
					if (pRow.RowState != DataRowState.Deleted && pRow["pos_art"].ToString() == currentArt)
					{
						decimal oldPrice = (dictOldPrice.ContainsKey(currentArt) ? dictOldPrice[currentArt] : 0m);
						decimal cost = (dictCost.ContainsKey(currentArt) ? dictCost[currentArt] : 0m);
						decimal margin = default(decimal);
						if (newPrice > 0m)
						{
							margin = Math.Round((newPrice - cost) / newPrice * 100m, 2);
						}
						pRow["pos_pvo"] = oldPrice;
						pRow["pos_pxc"] = cost;
						pRow["pos_mrg"] = margin;
						pRow["pos_prv"] = newPrice;
					}
				}
			}
			foreach (DataRow eRow in tEti.Rows)
			{
				if (eRow.RowState != DataRowState.Deleted && eRow["eti_art"].ToString() == currentArt)
				{
					eRow["eti_prv"] = newPrice;
					eRow["eti_pve"] = newPrice;
				}
			}
		}
		DataTable t2 = dSet.Tables["PosArt"];
		foreach (DataRow yy in t2.Rows)
		{
			tPos.ImportRow(yy);
		}
		dgv3.DataSource = tEti;
		lblVarCnt.Text = tEti.Rows.Count.ToString();
		if (sMsg != "")
		{
			MessageBox.Show(this, sMsg, "ERRORI");
		}
		lblVarCnt.Text = ((DataTable)dgv1.DataSource).Rows.Count.ToString();
		lblEtiCnt.Text = ((DataTable)dgv3.DataSource).Rows.Count.ToString();
		ClearArticleCache();
	}

	private void dgv1_CellClick(object sender, DataGridViewCellEventArgs e)
	{
		if (e.ColumnIndex != 0)
		{
			return;
		}
		CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
		if (cm.Position >= 0)
		{
			DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
			DataRow x = r.Row;
			switch ((string)x["pos_inv"])
			{
			case "0":
				x["pos_inv"] = "1";
				break;
			case "1":
				x["pos_inv"] = "3";
				break;
			case "3":
				x["pos_inv"] = "4";
				break;
			case "4":
				x["pos_inv"] = "0";
				break;
			}
		}
	}

	private void dgv1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
	{
		if (e.ColumnIndex == 3)
		{
			CurrencyManager cm = dgv1.BindingContext[dgv1.DataSource, dgv1.DataMember] as CurrencyManager;
			if (cm.Position >= 0)
			{
				DataRowView r = cm.List[dgv1.CurrentRow.Index] as DataRowView;
				DataRow x = r.Row;
				frmAnaArticolo f = new frmAnaArticolo();
				f._strArtCod = (string)x["pos_art"];
				f.ShowDialog();
				FillVarPos("", "");
				FillOfferte();
			}
		}
	}

	private void dgv2_CellClick(object sender, DataGridViewCellEventArgs e)
	{
	}

	private void dgv2_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
	{
		if (e.ColumnIndex != 2)
		{
			return;
		}
		CurrencyManager cm = dgv2.BindingContext[dgv2.DataSource, dgv2.DataMember] as CurrencyManager;
		if (cm.Position < 0)
		{
			return;
		}
		DataRowView r = cm.List[dgv2.CurrentRow.Index] as DataRowView;
		DataRow x = r.Row;
		frmGesOfferta f = new frmGesOfferta();
		f._strOftYea = (string)x["oft_yea"];
		f._strOftCod = (string)x["oft_cod"];
		f._strOftSta = (string)x["oft_sta"];
		f.ShowDialog(this);
		if (f._dasGen.Tables.IndexOf("GesOffTestate") >= 0)
		{
			DataTable t = f._dasGen.Tables["GesOffTestate"];
			if (t.Rows.Count > 0)
			{
				x["oft_des"] = t.Rows[0]["oft_des"];
				x["oft_dti"] = t.Rows[0]["oft_dti"];
				x["oft_dtf"] = t.Rows[0]["oft_dtf"];
				x["oft_sta"] = t.Rows[0]["oft_sta"];
				x["OftStd"] = t.Rows[0]["OftStd"];
			}
		}
	}

	private void dgv3_CellClick(object sender, DataGridViewCellEventArgs e)
	{
		if (e.ColumnIndex != 0)
		{
			return;
		}
		CurrencyManager cm = dgv3.BindingContext[dgv3.DataSource, dgv3.DataMember] as CurrencyManager;
		if (cm.Position >= 0)
		{
			DataRowView r = cm.List[dgv3.CurrentRow.Index] as DataRowView;
			DataRow x = r.Row;
			switch ((string)x["eti_inv"])
			{
			case "0":
				x["eti_inv"] = "1";
				break;
			case "1":
				x["eti_inv"] = "3";
				break;
			case "3":
				x["eti_inv"] = "4";
				break;
			case "4":
				x["eti_inv"] = "0";
				break;
			}
		}
		DataTable tEti = (DataTable)dgv3.DataSource;
		DataView v = new DataView(tEti, "eti_inv='0'", "", DataViewRowState.CurrentRows);
		lblEtiCnt.Text = v.Count.ToString();
	}

	private void dgv3_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
	{
		if (e.ColumnIndex == 2)
		{
			CurrencyManager cm = dgv3.BindingContext[dgv3.DataSource, dgv3.DataMember] as CurrencyManager;
			if (cm.Position >= 0)
			{
				DataRowView r = cm.List[dgv3.CurrentRow.Index] as DataRowView;
				DataRow x = r.Row;
				frmAnaArticolo f = new frmAnaArticolo();
				f._strArtCod = (string)x["eti_art"];
				f.ShowDialog();
				FillVarPos("", "");
				FillOfferte();
			}
		}
	}

	private void stampaEtichetteToolStripMenuItem_Click(object sender, EventArgs e)
	{
		dgv3.CommitEdit(DataGridViewDataErrorContexts.Commit);
		CurrencyManager cm = dgv3.BindingContext[dgv3.DataSource, dgv3.DataMember] as CurrencyManager;
		if (cm.Position >= 0)
		{
			DataRowView r = cm.List[dgv3.CurrentRow.Index] as DataRowView;
			DataRow x = r.Row;
			cm.EndCurrentEdit();
		}
	}

	private void refreshToolStripMenuItem_Click(object sender, EventArgs e)
	{
		try
		{
			int countP = 0, countE = 0;
			_clsVar.ConsolidaOfferteAllAvvio(out countP, out countE);
		}
		catch (Exception ex)
		{
			_clsFun.ErrorLog("frmGesVariazioni.refreshToolStripMenuItem_Click", ex.Message);
		}
		FillVarPos("", "");
		FillOfferte();
	}

	private void promoToolStripMenuItem_Click(object sender, EventArgs e)
	{
		new frmGesVarPromo().ShowDialog(this);
	}

	private void ImpEtiElettroniche(string strCod)
	{
		frmWait.ShowWait("Generazione ed esportazione etichette elettroniche in corso... ATTENDERE.");
		try
		{
			string s = "";
			string sMsg = "";
			ClearArticleCache();
			DataTable tPos = (DataTable)dgv1.DataSource;
			DataTable tTgr = _clsFun.FillTabSql("TabTipoGrammatura", "SELECT * FROM TabTipoGrammatura", bPrimo: false, _strConSql);
			DataTable tCal = _clsFun.FillTabSql("TabCalibro", "SELECT * FROM TabCalibro", bPrimo: false, _strConSql);
			DataTable tCat = _clsFun.FillTabSql("TabCategoria", "SELECT * FROM TabCategoria", bPrimo: false, _strConSql);
			DataTable tOri = _clsFun.FillTabSql("TabOrigine", "SELECT * FROM TabOrigine", bPrimo: false, _strConSql);
			DataTable tPro = _clsQry.LivPromo();
			s = "SELECT * FROM TabReparti WHERE tab_ann=0";
			DataTable tRep = _clsFun.FillTabSql("TabReparti", s, bPrimo: false, _strConSql);
			s = "SELECT * FROM TabIva WHERE tab_ann=0";
			DataTable tIva = _clsFun.FillTabSql("TabIva", s, bPrimo: false, _strConSql);
			DataTable tLne = _clsFun.FillTabSql("TabLisNegozi", "SELECT tab_lis FROM TabNegozi WHERE tab_tip = 'L'", bPrimo: false, _strConSql);
			string sLisNeg = ((tLne != null && tLne.Rows.Count > 0) ? ((string)tLne.Rows[0]["tab_lis"]) : "");
			DataTable tOffEti = _clsQry.OfsSeekEti(DateTime.Today.AddDays(-15.0));
			DataSet dSet = new DataSet();
			dSet.Tables.Add(tPos.Clone());
			dSet.Tables.Add(tTgr);
			dSet.Tables.Add(tCal);
			dSet.Tables.Add(tCat);
			dSet.Tables.Add(tOri);
			dSet.Tables.Add(tPro);
			DataTable t = new DataTable();
			t.Columns.Add("ord_art", typeof(string));
			t.Columns.Add("ord_ean", typeof(string));
			DataTable tDgv = ((dgv3 != null && dgv3.DataSource != null) ? (dgv3.DataSource as DataTable) : null);
			if (tDgv != null && tDgv.Rows.Count > 0)
			{
				HashSet<string> seenArt = new HashSet<string>();
				foreach (DataRow r in tDgv.Rows)
				{
					string code = "";
					if (tDgv.Columns.Contains("eti_art") && !DBNull.Value.Equals(r["eti_art"]))
					{
						code = r["eti_art"].ToString().Trim();
					}
					else if (tDgv.Columns.Contains("var_art") && !DBNull.Value.Equals(r["var_art"]))
					{
						code = r["var_art"].ToString().Trim();
					}
					if (!string.IsNullOrEmpty(code) && !seenArt.Contains(code))
					{
						seenArt.Add(code);
						DataRow nr = t.NewRow();
						nr["ord_art"] = code.PadLeft(7, '0');
						nr["ord_ean"] = (tDgv.Columns.Contains("eti_ean") ? Convert.ToString(r["eti_ean"]) : "");
						t.Rows.Add(nr);
					}
				}
			}
			if (t.Rows.Count == 0)
			{
				s = "SELECT * FROM GesVariazioni WHERE var_num='" + _clsDef.COD03Z + "' AND var_tip='" + _clsDef.VARETI + "'";
				DataTable tVarList = _clsFun.FillTabSql("GesVariazioni", s, bPrimo: false, _strConSql);
				HashSet<string> varArtSet = new HashSet<string>();
				foreach (DataRow r2 in tVarList.Rows)
				{
					if (r2["var_art"] != DBNull.Value && !string.IsNullOrEmpty(r2["var_art"].ToString().Trim()))
					{
						varArtSet.Add(r2["var_art"].ToString().Trim());
					}
				}
				if (varArtSet.Count > 0)
				{
					string inClause = string.Join(",", varArtSet.Select((string c) => "'" + c.Replace("'", "''") + "'"));
					s = "SELECT art_cod AS ord_art, '' AS ord_ean FROM AnaArticoli WHERE art_cod IN (" + inClause + ") ORDER BY art_cod";
				}
				else
				{
					s = "SELECT art_cod AS ord_art, '' AS ord_ean FROM AnaArticoli ORDER BY art_cod";
				}
				t = _clsFun.FillTabSql("AnaArticoli", s, bPrimo: false, _strConSql);
			}
			if (t.Rows.Count == 0)
			{
				frmWait.CloseWait();
				MessageBox.Show(this, "Nessun articolo presente nell'elenco etichette per l'esportazione.", "ESPORTAZIONE ETICHETTE ELETTRONICHE", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
				return;
			}
			DataTable tEti = ((DataTable)dgv3.DataSource).Clone();
			s = "SELECT * FROM GesVariazioni WHERE var_num='" + _clsDef.COD03Z + "'";
			DataTable tVar = _clsFun.FillTabSql("GesVariazioni", s, bPrimo: false, _strConSql);
			DataTable tPreloadVar = new DataTable();
			tPreloadVar.Columns.Add("var_art", typeof(string));
			foreach (DataRow y in t.Rows)
			{
				if (((string)y["ord_art"]).Trim() == "")
				{
					DataTable tTmp = _clsQry.ArtSeek((string)y["ord_ean"], "");
					if (tTmp.Rows.Count > 0)
					{
						y["ord_art"] = ((string)tTmp.Rows[0]["tmp_art"]).PadLeft(7, Convert.ToChar("0"));
						if ((string)tTmp.Rows[0]["tmp_sta"] == _clsDef.STANOA)
						{
							_clsQry.ArtAttiva((string)tTmp.Rows[0]["tmp_art"]);
						}
					}
				}
				if (((string)y["ord_art"]).Trim() != "")
				{
					DataRow nr2 = tPreloadVar.NewRow();
					nr2["var_art"] = ((string)y["ord_art"]).PadLeft(7, Convert.ToChar("0"));
					tPreloadVar.Rows.Add(nr2);
				}
			}
			PreloadArticlesCache(tPreloadVar, sLisNeg);
			HashSet<string> addedEti = new HashSet<string>();
			LookupCache cache = new LookupCache();
			foreach (DataRow r3 in tRep.Rows)
			{
				cache.Reparti[r3["tab_cod"].ToString()] = r3;
			}
			foreach (DataRow r4 in tIva.Rows)
			{
				cache.Iva[r4["tab_cod"].ToString()] = r4;
			}
			foreach (DataRow r5 in tTgr.Rows)
			{
				cache.Tgr[r5["tab_cod"].ToString()] = r5;
			}
			foreach (DataRow r6 in tCal.Rows)
			{
				cache.Cal[r6["tab_cod"].ToString()] = r6;
			}
			foreach (DataRow r7 in tCat.Rows)
			{
				cache.Cat[r7["tab_cod"].ToString()] = r7;
			}
			foreach (DataRow r8 in tOri.Rows)
			{
				cache.Ori[r8["tab_cod"].ToString()] = r8;
			}
			foreach (DataRow r9 in tOffEti.Rows)
			{
				cache.OffTes[r9["oft_cod"].ToString() + "|" + r9["ofa_art"].ToString()] = r9;
			}
			foreach (DataRow y2 in t.Rows)
			{
				if (((string)y2["ord_art"]).Trim() != "")
				{
					DataRow x = tVar.NewRow();
					x["var_tip"] = _clsDef.VARETI;
					x["var_num"] = _clsDef.COD03Z;
					x["var_art"] = y2["ord_art"];
					x["var_dti"] = DateTime.Today;
					x["var_dtv"] = DateTime.Today;
					FillEti(dSet, tEti, x, tOffEti, addedEti, cache);
				}
			}
			_clsVar.EtiElettroniche(strCod, tEti, null, null, this);
			FillVarEti();
			frmWait.CloseWait();
		}
		finally
		{
			frmWait.CloseWait();
			ClearArticleCache();
		}
	}

	private void cmbEtiRep_SelectionChangeCommitted(object sender, EventArgs e)
	{
		SalvaVar(_clsDef.VARETI);
		FillVarPos(_clsDef.VARETI, "");
		string sRep = "";
		if (cmbEtiRep.SelectedValue != null)
		{
			sRep = cmbEtiRep.SelectedValue.ToString();
		}
		if (sRep != "")
		{
			DataView v = new DataView((DataTable)dgv3.DataSource, "eti_rep='" + sRep + "'", "", DataViewRowState.CurrentRows);
			dgv3.DataSource = v.ToTable();
			lblEtiCnt.Text = v.Count.ToString();
		}
	}

	private void dgv1_CellContentClick(object sender, DataGridViewCellEventArgs e)
	{
	}

	private void dgv1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
	{
		if (e.RowIndex < 0 || e.RowIndex >= dgv1.Rows.Count)
		{
			return;
		}
		DataGridViewRow row = dgv1.Rows[e.RowIndex];
		DataRowView drv = row.DataBoundItem as DataRowView;
		if (drv != null && drv.Row != null && drv.Row.RowState != DataRowState.Deleted)
		{
			decimal pos_prv = ((drv.Row["pos_prv"] != DBNull.Value) ? Convert.ToDecimal(drv.Row["pos_prv"]) : 0m);
			decimal pos_pvo = ((drv.Row["pos_pvo"] != DBNull.Value) ? Convert.ToDecimal(drv.Row["pos_pvo"]) : 0m);
			if (pos_pvo > 0m)
			{
				if (pos_prv > pos_pvo)
				{
					e.CellStyle.ForeColor = Color.Blue;
				}
				else if (pos_prv < pos_pvo)
				{
					e.CellStyle.ForeColor = Color.Red;
				}
			}
		}
	}

	private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
	{
	}

}
}

