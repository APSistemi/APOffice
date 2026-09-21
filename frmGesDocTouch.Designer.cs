namespace APOffice
{
    partial class frmGesDocTouch
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle3 = new System.Windows.Forms.DataGridViewCellStyle();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.esciToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.letturaDaScontrinoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.utilitàToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.eliminaRigheToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripSeparator();
            this.eliminaRigheCancellateToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.iminaDocumentoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.importDaDocumentiToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.lottiToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dgv1 = new APOffice.APDataGridView();
            this.dgv2 = new APOffice.APDataGridView();
            this.panel1 = new System.Windows.Forms.Panel();
            this.btnRowDbl = new System.Windows.Forms.Button();
            this.btnArtMod = new System.Windows.Forms.Button();
            this.btnArtNew = new System.Windows.Forms.Button();
            this.btnEtiArt = new System.Windows.Forms.Button();
            this.btnEtiIng = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.lblMftNum = new System.Windows.Forms.Label();
            this.lblMftYea = new System.Windows.Forms.Label();
            this.cmbMftTdc = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.cmbMftCfo = new System.Windows.Forms.ComboBox();
            this.btnCfo = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.dtpMftDdt = new System.Windows.Forms.DateTimePicker();
            this.txtMftNdo = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtMftNo1 = new System.Windows.Forms.TextBox();
            this.btnPrn = new System.Windows.Forms.Button();
            this.cmbMftTpg = new System.Windows.Forms.ComboBox();
            this.label14 = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.label16 = new System.Windows.Forms.Label();
            this.label17 = new System.Windows.Forms.Label();
            this.lblTotImp = new System.Windows.Forms.Label();
            this.lblTotIva = new System.Windows.Forms.Label();
            this.lblTotTot = new System.Windows.Forms.Label();
            this.cmbMftSta = new System.Windows.Forms.ComboBox();
            this.label18 = new System.Windows.Forms.Label();
            this.label19 = new System.Windows.Forms.Label();
            this.txtMftSco = new System.Windows.Forms.TextBox();
            this.cmbMftSco = new System.Windows.Forms.ComboBox();
            this.btnMftSco = new System.Windows.Forms.Button();
            this.label21 = new System.Windows.Forms.Label();
            this.cmbMftNeg = new System.Windows.Forms.ComboBox();
            this.label22 = new System.Windows.Forms.Label();
            this.lblTotQta = new System.Windows.Forms.Label();
            this.label23 = new System.Windows.Forms.Label();
            this.cmbMftDes = new System.Windows.Forms.ComboBox();
            this.btnArtIns = new System.Windows.Forms.Button();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv2)).BeginInit();
            this.panel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem,
            this.letturaDaScontrinoToolStripMenuItem,
            this.utilitàToolStripMenuItem,
            this.importDaDocumentiToolStripMenuItem,
            this.lottiToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(951, 24);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // esciToolStripMenuItem
            // 
            this.esciToolStripMenuItem.Name = "esciToolStripMenuItem";
            this.esciToolStripMenuItem.Size = new System.Drawing.Size(39, 20);
            this.esciToolStripMenuItem.Text = "Esci";
            this.esciToolStripMenuItem.Click += new System.EventHandler(this.esciToolStripMenuItem_Click);
            // 
            // letturaDaScontrinoToolStripMenuItem
            // 
            this.letturaDaScontrinoToolStripMenuItem.Name = "letturaDaScontrinoToolStripMenuItem";
            this.letturaDaScontrinoToolStripMenuItem.Size = new System.Drawing.Size(125, 20);
            this.letturaDaScontrinoToolStripMenuItem.Text = "Lettura da scontrino";
            this.letturaDaScontrinoToolStripMenuItem.Click += new System.EventHandler(this.letturaDaScontrinoToolStripMenuItem_Click);
            // 
            // utilitàToolStripMenuItem
            // 
            this.utilitàToolStripMenuItem.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.utilitàToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.eliminaRigheToolStripMenuItem,
            this.toolStripMenuItem1,
            this.eliminaRigheCancellateToolStripMenuItem,
            this.toolStripSeparator1,
            this.iminaDocumentoToolStripMenuItem});
            this.utilitàToolStripMenuItem.Name = "utilitàToolStripMenuItem";
            this.utilitàToolStripMenuItem.Size = new System.Drawing.Size(50, 20);
            this.utilitàToolStripMenuItem.Text = "Utilità";
            // 
            // eliminaRigheToolStripMenuItem
            // 
            this.eliminaRigheToolStripMenuItem.Name = "eliminaRigheToolStripMenuItem";
            this.eliminaRigheToolStripMenuItem.Size = new System.Drawing.Size(255, 22);
            this.eliminaRigheToolStripMenuItem.Text = "Elimina tutte righe del documento";
            this.eliminaRigheToolStripMenuItem.Click += new System.EventHandler(this.eliminaRigheToolStripMenuItem_Click);
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(252, 6);
            // 
            // eliminaRigheCancellateToolStripMenuItem
            // 
            this.eliminaRigheCancellateToolStripMenuItem.Name = "eliminaRigheCancellateToolStripMenuItem";
            this.eliminaRigheCancellateToolStripMenuItem.Size = new System.Drawing.Size(255, 22);
            this.eliminaRigheCancellateToolStripMenuItem.Text = "Elimina righe cancellate";
            this.eliminaRigheCancellateToolStripMenuItem.Click += new System.EventHandler(this.eliminaRigheCancellateToolStripMenuItem_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(252, 6);
            // 
            // iminaDocumentoToolStripMenuItem
            // 
            this.iminaDocumentoToolStripMenuItem.Name = "iminaDocumentoToolStripMenuItem";
            this.iminaDocumentoToolStripMenuItem.Size = new System.Drawing.Size(255, 22);
            this.iminaDocumentoToolStripMenuItem.Text = "Eliimina documento";
            this.iminaDocumentoToolStripMenuItem.Click += new System.EventHandler(this.iminaDocumentoToolStripMenuItem_Click);
            // 
            // importDaDocumentiToolStripMenuItem
            // 
            this.importDaDocumentiToolStripMenuItem.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.importDaDocumentiToolStripMenuItem.Name = "importDaDocumentiToolStripMenuItem";
            this.importDaDocumentiToolStripMenuItem.Size = new System.Drawing.Size(132, 20);
            this.importDaDocumentiToolStripMenuItem.Text = "Import da documenti";
            this.importDaDocumentiToolStripMenuItem.Click += new System.EventHandler(this.importDaDocumentiToolStripMenuItem_Click);
            // 
            // lottiToolStripMenuItem
            // 
            this.lottiToolStripMenuItem.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.lottiToolStripMenuItem.Name = "lottiToolStripMenuItem";
            this.lottiToolStripMenuItem.Size = new System.Drawing.Size(43, 20);
            this.lottiToolStripMenuItem.Text = "Lotti";
            this.lottiToolStripMenuItem.Click += new System.EventHandler(this.lottiToolStripMenuItem_Click);
            // 
            // dgv1
            // 
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.Lavender;
            this.dgv1.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgv1.DefaultCellStyle = dataGridViewCellStyle2;
            this.dgv1.Location = new System.Drawing.Point(7, 253);
            this.dgv1.Name = "dgv1";
            this.dgv1.RowHeadersWidth = 20;
            this.dgv1.Size = new System.Drawing.Size(935, 280);
            this.dgv1.TabIndex = 1;
            this.dgv1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv1_CellContentClick);
            this.dgv1.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv1_CellDoubleClick);
            this.dgv1.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv1_CellEndEdit);
            this.dgv1.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv1_CellValueChanged);
            this.dgv1.CurrentCellChanged += new System.EventHandler(this.dgv1_CurrentCellChanged);
            this.dgv1.CurrentCellDirtyStateChanged += new System.EventHandler(this.dgv1_CurrentCellDirtyStateChanged);
            this.dgv1.EditingControlShowing += new System.Windows.Forms.DataGridViewEditingControlShowingEventHandler(this.dgv1_EditingControlShowing);
            // 
            // dgv2
            // 
            dataGridViewCellStyle3.BackColor = System.Drawing.Color.Lavender;
            this.dgv2.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle3;
            this.dgv2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv2.Location = new System.Drawing.Point(7, 539);
            this.dgv2.Name = "dgv2";
            this.dgv2.RowHeadersWidth = 20;
            this.dgv2.Size = new System.Drawing.Size(412, 87);
            this.dgv2.TabIndex = 2;
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.btnArtIns);
            this.panel1.Controls.Add(this.btnRowDbl);
            this.panel1.Controls.Add(this.btnArtMod);
            this.panel1.Controls.Add(this.btnArtNew);
            this.panel1.Controls.Add(this.btnEtiArt);
            this.panel1.Controls.Add(this.btnEtiIng);
            this.panel1.Location = new System.Drawing.Point(7, 189);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(939, 67);
            this.panel1.TabIndex = 3;
            // 
            // btnRowDbl
            // 
            this.btnRowDbl.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnRowDbl.Location = new System.Drawing.Point(306, 5);
            this.btnRowDbl.Name = "btnRowDbl";
            this.btnRowDbl.Size = new System.Drawing.Size(147, 51);
            this.btnRowDbl.TabIndex = 27;
            this.btnRowDbl.Text = "Duplica riga prodotto";
            this.btnRowDbl.UseVisualStyleBackColor = true;
            this.btnRowDbl.Click += new System.EventHandler(this.btnRowDbl_Click);
            // 
            // btnArtMod
            // 
            this.btnArtMod.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnArtMod.Location = new System.Drawing.Point(158, 5);
            this.btnArtMod.Name = "btnArtMod";
            this.btnArtMod.Size = new System.Drawing.Size(142, 51);
            this.btnArtMod.TabIndex = 26;
            this.btnArtMod.Text = "Modifica riga prodotto";
            this.btnArtMod.UseVisualStyleBackColor = true;
            this.btnArtMod.Click += new System.EventHandler(this.btnArtMod_Click);
            // 
            // btnArtNew
            // 
            this.btnArtNew.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnArtNew.Location = new System.Drawing.Point(5, 5);
            this.btnArtNew.Name = "btnArtNew";
            this.btnArtNew.Size = new System.Drawing.Size(148, 51);
            this.btnArtNew.TabIndex = 25;
            this.btnArtNew.Text = "Inserisci prodotto pesato";
            this.btnArtNew.UseVisualStyleBackColor = true;
            this.btnArtNew.Click += new System.EventHandler(this.btnArtNew_Click);
            // 
            // btnEtiArt
            // 
            this.btnEtiArt.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnEtiArt.Location = new System.Drawing.Point(778, 5);
            this.btnEtiArt.Name = "btnEtiArt";
            this.btnEtiArt.Size = new System.Drawing.Size(155, 51);
            this.btnEtiArt.TabIndex = 24;
            this.btnEtiArt.Text = "Etichetta riepilogo articolo";
            this.btnEtiArt.UseVisualStyleBackColor = true;
            this.btnEtiArt.Click += new System.EventHandler(this.btnEtiArt_Click);
            // 
            // btnEtiIng
            // 
            this.btnEtiIng.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnEtiIng.Location = new System.Drawing.Point(618, 5);
            this.btnEtiIng.Name = "btnEtiIng";
            this.btnEtiIng.Size = new System.Drawing.Size(155, 51);
            this.btnEtiIng.TabIndex = 23;
            this.btnEtiIng.Text = "Etichetta ingredienti singola riga";
            this.btnEtiIng.UseVisualStyleBackColor = true;
            this.btnEtiIng.Click += new System.EventHandler(this.btnEtiIng_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(8, 40);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(59, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "Movimento";
            // 
            // lblMftNum
            // 
            this.lblMftNum.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblMftNum.Location = new System.Drawing.Point(123, 34);
            this.lblMftNum.Name = "lblMftNum";
            this.lblMftNum.Size = new System.Drawing.Size(63, 17);
            this.lblMftNum.TabIndex = 5;
            this.lblMftNum.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lblMftYea
            // 
            this.lblMftYea.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblMftYea.Location = new System.Drawing.Point(78, 35);
            this.lblMftYea.Name = "lblMftYea";
            this.lblMftYea.Size = new System.Drawing.Size(36, 17);
            this.lblMftYea.TabIndex = 6;
            this.lblMftYea.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // cmbMftTdc
            // 
            this.cmbMftTdc.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbMftTdc.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbMftTdc.FormattingEnabled = true;
            this.cmbMftTdc.Location = new System.Drawing.Point(197, 32);
            this.cmbMftTdc.Name = "cmbMftTdc";
            this.cmbMftTdc.Size = new System.Drawing.Size(233, 28);
            this.cmbMftTdc.TabIndex = 8;
            this.cmbMftTdc.SelectionChangeCommitted += new System.EventHandler(this.cmbMftTdc_SelectionChangeCommitted);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(436, 39);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(63, 13);
            this.label3.TabIndex = 9;
            this.label3.Text = "Destinatario";
            // 
            // cmbMftCfo
            // 
            this.cmbMftCfo.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbMftCfo.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbMftCfo.FormattingEnabled = true;
            this.cmbMftCfo.Location = new System.Drawing.Point(505, 31);
            this.cmbMftCfo.Name = "cmbMftCfo";
            this.cmbMftCfo.Size = new System.Drawing.Size(338, 28);
            this.cmbMftCfo.TabIndex = 10;
            this.cmbMftCfo.SelectionChangeCommitted += new System.EventHandler(this.cmbMftCfo_SelectionChangeCommitted);
            // 
            // btnCfo
            // 
            this.btnCfo.Location = new System.Drawing.Point(849, 30);
            this.btnCfo.Name = "btnCfo";
            this.btnCfo.Size = new System.Drawing.Size(30, 29);
            this.btnCfo.TabIndex = 11;
            this.btnCfo.Text = "...";
            this.btnCfo.UseVisualStyleBackColor = true;
            this.btnCfo.Click += new System.EventHandler(this.btnCfo_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(10, 65);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(92, 20);
            this.label4.TabIndex = 12;
            this.label4.Text = "Documento";
            // 
            // dtpMftDdt
            // 
            this.dtpMftDdt.CalendarFont = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dtpMftDdt.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dtpMftDdt.Location = new System.Drawing.Point(123, 88);
            this.dtpMftDdt.Name = "dtpMftDdt";
            this.dtpMftDdt.Size = new System.Drawing.Size(309, 31);
            this.dtpMftDdt.TabIndex = 13;
            // 
            // txtMftNdo
            // 
            this.txtMftNdo.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtMftNdo.Location = new System.Drawing.Point(7, 88);
            this.txtMftNdo.MaxLength = 10;
            this.txtMftNdo.Name = "txtMftNdo";
            this.txtMftNdo.Size = new System.Drawing.Size(110, 31);
            this.txtMftNdo.TabIndex = 14;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(465, 131);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(30, 13);
            this.label2.TabIndex = 15;
            this.label2.Text = "Note";
            // 
            // txtMftNo1
            // 
            this.txtMftNo1.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtMftNo1.Location = new System.Drawing.Point(466, 148);
            this.txtMftNo1.MaxLength = 100;
            this.txtMftNo1.Name = "txtMftNo1";
            this.txtMftNo1.Size = new System.Drawing.Size(476, 31);
            this.txtMftNo1.TabIndex = 16;
            // 
            // btnPrn
            // 
            this.btnPrn.Location = new System.Drawing.Point(880, 30);
            this.btnPrn.Name = "btnPrn";
            this.btnPrn.Size = new System.Drawing.Size(62, 29);
            this.btnPrn.TabIndex = 17;
            this.btnPrn.Text = "Stampa";
            this.btnPrn.UseVisualStyleBackColor = true;
            this.btnPrn.Click += new System.EventHandler(this.btnPrn_Click);
            // 
            // cmbMftTpg
            // 
            this.cmbMftTpg.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbMftTpg.FormattingEnabled = true;
            this.cmbMftTpg.Location = new System.Drawing.Point(7, 148);
            this.cmbMftTpg.Name = "cmbMftTpg";
            this.cmbMftTpg.Size = new System.Drawing.Size(132, 33);
            this.cmbMftTpg.TabIndex = 20;
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(5, 132);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(70, 13);
            this.label14.TabIndex = 21;
            this.label14.Text = "T.pagamento";
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(663, 585);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(93, 13);
            this.label15.TabIndex = 22;
            this.label15.Text = "Totale documento";
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(663, 549);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(54, 13);
            this.label16.TabIndex = 23;
            this.label16.Text = "Imponibile";
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(663, 576);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(44, 13);
            this.label17.TabIndex = 24;
            this.label17.Text = "Imposta";
            // 
            // lblTotImp
            // 
            this.lblTotImp.BackColor = System.Drawing.SystemColors.ControlLight;
            this.lblTotImp.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblTotImp.Location = new System.Drawing.Point(789, 539);
            this.lblTotImp.Name = "lblTotImp";
            this.lblTotImp.Size = new System.Drawing.Size(100, 23);
            this.lblTotImp.TabIndex = 25;
            this.lblTotImp.Text = "0";
            this.lblTotImp.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblTotIva
            // 
            this.lblTotIva.BackColor = System.Drawing.SystemColors.ControlLight;
            this.lblTotIva.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblTotIva.Location = new System.Drawing.Point(789, 566);
            this.lblTotIva.Name = "lblTotIva";
            this.lblTotIva.Size = new System.Drawing.Size(100, 23);
            this.lblTotIva.TabIndex = 26;
            this.lblTotIva.Text = "0";
            this.lblTotIva.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblTotTot
            // 
            this.lblTotTot.BackColor = System.Drawing.SystemColors.ControlLight;
            this.lblTotTot.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblTotTot.Location = new System.Drawing.Point(789, 592);
            this.lblTotTot.Name = "lblTotTot";
            this.lblTotTot.Size = new System.Drawing.Size(100, 23);
            this.lblTotTot.TabIndex = 27;
            this.lblTotTot.Text = "0";
            this.lblTotTot.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // cmbMftSta
            // 
            this.cmbMftSta.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbMftSta.FormattingEnabled = true;
            this.cmbMftSta.Location = new System.Drawing.Point(280, 146);
            this.cmbMftSta.Name = "cmbMftSta";
            this.cmbMftSta.Size = new System.Drawing.Size(182, 33);
            this.cmbMftSta.TabIndex = 29;
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Location = new System.Drawing.Point(277, 130);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(32, 13);
            this.label18.TabIndex = 30;
            this.label18.Text = "Stato";
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Location = new System.Drawing.Point(430, 571);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(74, 13);
            this.label19.TabIndex = 31;
            this.label19.Text = "Sconto fattura";
            // 
            // txtMftSco
            // 
            this.txtMftSco.Location = new System.Drawing.Point(499, 587);
            this.txtMftSco.Name = "txtMftSco";
            this.txtMftSco.Size = new System.Drawing.Size(86, 20);
            this.txtMftSco.TabIndex = 32;
            // 
            // cmbMftSco
            // 
            this.cmbMftSco.Font = new System.Drawing.Font("Microsoft Sans Serif", 6.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbMftSco.FormattingEnabled = true;
            this.cmbMftSco.Location = new System.Drawing.Point(433, 587);
            this.cmbMftSco.Name = "cmbMftSco";
            this.cmbMftSco.Size = new System.Drawing.Size(57, 20);
            this.cmbMftSco.TabIndex = 33;
            // 
            // btnMftSco
            // 
            this.btnMftSco.Location = new System.Drawing.Point(591, 585);
            this.btnMftSco.Name = "btnMftSco";
            this.btnMftSco.Size = new System.Drawing.Size(36, 23);
            this.btnMftSco.TabIndex = 34;
            this.btnMftSco.Text = "Calcola";
            this.btnMftSco.UseVisualStyleBackColor = true;
            this.btnMftSco.Click += new System.EventHandler(this.btnMftSco_Click);
            // 
            // label21
            // 
            this.label21.AutoSize = true;
            this.label21.Location = new System.Drawing.Point(143, 132);
            this.label21.Name = "label21";
            this.label21.Size = new System.Drawing.Size(46, 13);
            this.label21.TabIndex = 35;
            this.label21.Text = "Negozio";
            // 
            // cmbMftNeg
            // 
            this.cmbMftNeg.Font = new System.Drawing.Font("Microsoft Sans Serif", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbMftNeg.FormattingEnabled = true;
            this.cmbMftNeg.Location = new System.Drawing.Point(145, 146);
            this.cmbMftNeg.Name = "cmbMftNeg";
            this.cmbMftNeg.Size = new System.Drawing.Size(129, 33);
            this.cmbMftNeg.TabIndex = 36;
            // 
            // label22
            // 
            this.label22.AutoSize = true;
            this.label22.Location = new System.Drawing.Point(430, 546);
            this.label22.Name = "label22";
            this.label22.Size = new System.Drawing.Size(32, 13);
            this.label22.TabIndex = 37;
            this.label22.Text = "Pezzi";
            // 
            // lblTotQta
            // 
            this.lblTotQta.BackColor = System.Drawing.SystemColors.ControlLight;
            this.lblTotQta.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.lblTotQta.Location = new System.Drawing.Point(499, 539);
            this.lblTotQta.Name = "lblTotQta";
            this.lblTotQta.Size = new System.Drawing.Size(86, 23);
            this.lblTotQta.TabIndex = 38;
            this.lblTotQta.Text = "0";
            this.lblTotQta.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label23
            // 
            this.label23.AutoSize = true;
            this.label23.Location = new System.Drawing.Point(436, 68);
            this.label23.Name = "label23";
            this.label23.Size = new System.Drawing.Size(68, 13);
            this.label23.TabIndex = 39;
            this.label23.Text = "Destinazione";
            // 
            // cmbMftDes
            // 
            this.cmbMftDes.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbMftDes.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbMftDes.FormattingEnabled = true;
            this.cmbMftDes.Location = new System.Drawing.Point(438, 88);
            this.cmbMftDes.Name = "cmbMftDes";
            this.cmbMftDes.Size = new System.Drawing.Size(500, 28);
            this.cmbMftDes.TabIndex = 40;
            // 
            // btnArtIns
            // 
            this.btnArtIns.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnArtIns.Location = new System.Drawing.Point(459, 5);
            this.btnArtIns.Name = "btnArtIns";
            this.btnArtIns.Size = new System.Drawing.Size(155, 51);
            this.btnArtIns.TabIndex = 28;
            this.btnArtIns.Text = "Inserisci prodotto";
            this.btnArtIns.UseVisualStyleBackColor = true;
            this.btnArtIns.Click += new System.EventHandler(this.btnArtIns_Click);
            // 
            // frmGesDocTouch
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(951, 631);
            this.ControlBox = false;
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.cmbMftDes);
            this.Controls.Add(this.label23);
            this.Controls.Add(this.lblTotQta);
            this.Controls.Add(this.label22);
            this.Controls.Add(this.cmbMftNeg);
            this.Controls.Add(this.label21);
            this.Controls.Add(this.btnMftSco);
            this.Controls.Add(this.cmbMftSco);
            this.Controls.Add(this.txtMftSco);
            this.Controls.Add(this.label19);
            this.Controls.Add(this.label18);
            this.Controls.Add(this.cmbMftSta);
            this.Controls.Add(this.lblTotTot);
            this.Controls.Add(this.lblTotIva);
            this.Controls.Add(this.lblTotImp);
            this.Controls.Add(this.label17);
            this.Controls.Add(this.label16);
            this.Controls.Add(this.label15);
            this.Controls.Add(this.cmbMftTpg);
            this.Controls.Add(this.label14);
            this.Controls.Add(this.btnPrn);
            this.Controls.Add(this.txtMftNo1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtMftNdo);
            this.Controls.Add(this.dtpMftDdt);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.btnCfo);
            this.Controls.Add(this.cmbMftCfo);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.cmbMftTdc);
            this.Controls.Add(this.lblMftYea);
            this.Controls.Add(this.lblMftNum);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.dgv2);
            this.Controls.Add(this.menuStrip1);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmGesDocTouch";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Gestione documento";
            this.Load += new System.EventHandler(this.frmGesDocumento_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmGesDocumento_KeyDown);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv2)).EndInit();
            this.panel1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem esciToolStripMenuItem;
        private APOffice.APDataGridView dgv1;
        private APOffice.APDataGridView dgv2;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label lblMftNum;
        private System.Windows.Forms.Label lblMftYea;
        private System.Windows.Forms.ComboBox cmbMftTdc;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cmbMftCfo;
        private System.Windows.Forms.Button btnCfo;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.DateTimePicker dtpMftDdt;
        private System.Windows.Forms.TextBox txtMftNdo;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtMftNo1;
        private System.Windows.Forms.Button btnPrn;
        private System.Windows.Forms.ComboBox cmbMftTpg;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.Label lblTotImp;
        private System.Windows.Forms.Label lblTotIva;
        private System.Windows.Forms.Label lblTotTot;
        private System.Windows.Forms.ToolStripMenuItem letturaDaScontrinoToolStripMenuItem;
        private System.Windows.Forms.ComboBox cmbMftSta;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.Label label19;
        private System.Windows.Forms.TextBox txtMftSco;
        private System.Windows.Forms.ComboBox cmbMftSco;
        private System.Windows.Forms.Button btnMftSco;
        private System.Windows.Forms.ToolStripMenuItem utilitàToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem eliminaRigheToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem importDaDocumentiToolStripMenuItem;
        private System.Windows.Forms.Label label21;
        private System.Windows.Forms.ComboBox cmbMftNeg;
        private System.Windows.Forms.Label label22;
        private System.Windows.Forms.Label lblTotQta;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem eliminaRigheCancellateToolStripMenuItem;
        private System.Windows.Forms.Label label23;
        private System.Windows.Forms.ComboBox cmbMftDes;
        private System.Windows.Forms.Button btnEtiArt;
        private System.Windows.Forms.Button btnEtiIng;
        private System.Windows.Forms.Button btnArtMod;
        private System.Windows.Forms.Button btnArtNew;
        private System.Windows.Forms.Button btnRowDbl;
        private System.Windows.Forms.ToolStripMenuItem lottiToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem iminaDocumentoToolStripMenuItem;
        private System.Windows.Forms.Button btnArtIns;
    }
}
