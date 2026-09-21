namespace APOffice
{
    partial class frmGesOffTransazione
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.esciToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dgv1 = new APOffice.APDataGridView();
            this.btnNew = new System.Windows.Forms.Button();
            this.btnInv = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.btnDayNew = new System.Windows.Forms.Button();
            this.label17 = new System.Windows.Forms.Label();
            this.dgv2 = new APOffice.APDataGridView();
            this.txtOtrRen = new System.Windows.Forms.TextBox();
            this.label16 = new System.Windows.Forms.Label();
            this.txtOtrRep = new System.Windows.Forms.TextBox();
            this.label15 = new System.Windows.Forms.Label();
            this.btnOtrMix = new System.Windows.Forms.Button();
            this.chkOtrVof = new System.Windows.Forms.CheckBox();
            this.chkOtrAnn = new System.Windows.Forms.CheckBox();
            this.dtpOtrDtf = new System.Windows.Forms.DateTimePicker();
            this.dtpOtrDti = new System.Windows.Forms.DateTimePicker();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.txtOtrMix = new System.Windows.Forms.TextBox();
            this.label14 = new System.Windows.Forms.Label();
            this.txtOtrSmx = new System.Windows.Forms.TextBox();
            this.label13 = new System.Windows.Forms.Label();
            this.txtOtrSmi = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.txtOtrSgl = new System.Windows.Forms.TextBox();
            this.label11 = new System.Windows.Forms.Label();
            this.txtOtrPas = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.txtOtrVal = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.cmbOtrTip = new System.Windows.Forms.ComboBox();
            this.cmbOtrCam = new System.Windows.Forms.ComboBox();
            this.cmbOtrNeg = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.cmbOtrGru = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtOtrDes = new System.Windows.Forms.TextBox();
            this.txtOtrCod = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.cmbOtrTca = new System.Windows.Forms.ComboBox();
            this.label18 = new System.Windows.Forms.Label();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv2)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(826, 24);
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
            // dgv1
            // 
            this.dgv1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.Lavender;
            this.dgv1.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle2;
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv1.Location = new System.Drawing.Point(12, 84);
            this.dgv1.Name = "dgv1";
            this.dgv1.RowHeadersWidth = 20;
            this.dgv1.Size = new System.Drawing.Size(840, 240);
            this.dgv1.TabIndex = 1;
            this.dgv1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv1_CellClick);
            this.dgv1.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv1_CellContentClick);
            this.dgv1.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv1_CellEndEdit);
            this.dgv1.CurrentCellChanged += new System.EventHandler(this.dgv1_CurrentCellChanged);
            this.dgv1.EditingControlShowing += new System.Windows.Forms.DataGridViewEditingControlShowingEventHandler(this.dgv1_EditingControlShowing);
            // 
            // btnNew
            // 
            this.btnNew.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnNew.Location = new System.Drawing.Point(12, 34);
            this.btnNew.Name = "btnNew";
            this.btnNew.Size = new System.Drawing.Size(160, 42);
            this.btnNew.TabIndex = 2;
            this.btnNew.Text = "Nuova offerta";
            this.btnNew.UseVisualStyleBackColor = true;
            this.btnNew.Click += new System.EventHandler(this.btnNew_Click);
            // 
            // btnInv
            // 
            this.btnInv.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnInv.Font = new System.Drawing.Font("Segoe UI", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnInv.Location = new System.Drawing.Point(682, 34);
            this.btnInv.Name = "btnInv";
            this.btnInv.Size = new System.Drawing.Size(170, 42);
            this.btnInv.TabIndex = 3;
            this.btnInv.Text = "Invio a cassa";
            this.btnInv.UseVisualStyleBackColor = true;
            this.btnInv.Click += new System.EventHandler(this.btnInv_Click);
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.BackColor = System.Drawing.Color.AliceBlue;
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.label18);
            this.panel1.Controls.Add(this.cmbOtrTca);
            this.panel1.Controls.Add(this.btnDayNew);
            this.panel1.Controls.Add(this.label17);
            this.panel1.Controls.Add(this.dgv2);
            this.panel1.Controls.Add(this.txtOtrRen);
            this.panel1.Controls.Add(this.label16);
            this.panel1.Controls.Add(this.txtOtrRep);
            this.panel1.Controls.Add(this.label15);
            this.panel1.Controls.Add(this.btnOtrMix);
            this.panel1.Controls.Add(this.chkOtrVof);
            this.panel1.Controls.Add(this.chkOtrAnn);
            this.panel1.Controls.Add(this.dtpOtrDtf);
            this.panel1.Controls.Add(this.dtpOtrDti);
            this.panel1.Controls.Add(this.label8);
            this.panel1.Controls.Add(this.label7);
            this.panel1.Controls.Add(this.txtOtrMix);
            this.panel1.Controls.Add(this.label14);
            this.panel1.Controls.Add(this.txtOtrSmx);
            this.panel1.Controls.Add(this.label13);
            this.panel1.Controls.Add(this.txtOtrSmi);
            this.panel1.Controls.Add(this.label12);
            this.panel1.Controls.Add(this.txtOtrSgl);
            this.panel1.Controls.Add(this.label11);
            this.panel1.Controls.Add(this.txtOtrPas);
            this.panel1.Controls.Add(this.label10);
            this.panel1.Controls.Add(this.txtOtrVal);
            this.panel1.Controls.Add(this.label9);
            this.panel1.Controls.Add(this.label6);
            this.panel1.Controls.Add(this.cmbOtrTip);
            this.panel1.Controls.Add(this.cmbOtrCam);
            this.panel1.Controls.Add(this.cmbOtrNeg);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.cmbOtrGru);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.txtOtrDes);
            this.panel1.Controls.Add(this.txtOtrCod);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.panel1.Location = new System.Drawing.Point(12, 332);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(840, 270);
            this.panel1.TabIndex = 4;
            // 
            // btnDayNew
            // 
            this.btnDayNew.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnDayNew.Location = new System.Drawing.Point(744, 4);
            this.btnDayNew.Name = "btnDayNew";
            this.btnDayNew.Size = new System.Drawing.Size(88, 24);
            this.btnDayNew.TabIndex = 44;
            this.btnDayNew.Text = "Nuova";
            this.btnDayNew.UseVisualStyleBackColor = true;
            this.btnDayNew.Click += new System.EventHandler(this.btnDayNew_Click);
            // 
            // label17
            // 
            this.label17.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.label17.AutoSize = true;
            this.label17.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label17.Location = new System.Drawing.Point(600, 9);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(68, 15);
            this.label17.TabIndex = 43;
            this.label17.Text = "Date attive";
            // 
            // dgv2
            // 
            this.dgv2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgv2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv2.Location = new System.Drawing.Point(600, 31);
            this.dgv2.Name = "dgv2";
            this.dgv2.RowHeadersWidth = 20;
            this.dgv2.Size = new System.Drawing.Size(232, 228);
            this.dgv2.TabIndex = 42;
            this.dgv2.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv2_CellClick);
            // 
            // txtOtrRen
            // 
            this.txtOtrRen.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtOtrRen.Location = new System.Drawing.Point(143, 183);
            this.txtOtrRen.MaxLength = 50;
            this.txtOtrRen.Name = "txtOtrRen";
            this.txtOtrRen.Size = new System.Drawing.Size(452, 21);
            this.txtOtrRen.TabIndex = 41;
            this.txtOtrRen.TextChanged += new System.EventHandler(this.txtOtr_TextChanged);
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(3, 185);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(128, 13);
            this.label16.TabIndex = 40;
            this.label16.Text = "Filtro reparti esclusi (sep .)";
            // 
            // txtOtrRep
            // 
            this.txtOtrRep.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtOtrRep.Location = new System.Drawing.Point(143, 154);
            this.txtOtrRep.MaxLength = 50;
            this.txtOtrRep.Name = "txtOtrRep";
            this.txtOtrRep.Size = new System.Drawing.Size(452, 21);
            this.txtOtrRep.TabIndex = 39;
            this.txtOtrRep.TextChanged += new System.EventHandler(this.txtOtr_TextChanged);
            this.txtOtrRep.Leave += new System.EventHandler(this.txtOtrDes_Leave);
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(2, 159);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(125, 13);
            this.label15.TabIndex = 38;
            this.label15.Text = "Filtro reparti inclusi (sep .)";
            // 
            // btnOtrMix
            // 
            this.btnOtrMix.Location = new System.Drawing.Point(561, 114);
            this.btnOtrMix.Name = "btnOtrMix";
            this.btnOtrMix.Size = new System.Drawing.Size(34, 23);
            this.btnOtrMix.TabIndex = 37;
            this.btnOtrMix.Text = "MIX";
            this.btnOtrMix.UseVisualStyleBackColor = true;
            this.btnOtrMix.Click += new System.EventHandler(this.btnOtrMix_Click);
            // 
            // chkOtrVof
            // 
            this.chkOtrVof.AutoSize = true;
            this.chkOtrVof.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkOtrVof.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkOtrVof.Location = new System.Drawing.Point(6, 115);
            this.chkOtrVof.Name = "chkOtrVof";
            this.chkOtrVof.Size = new System.Drawing.Size(156, 19);
            this.chkOtrVof.TabIndex = 36;
            this.chkOtrVof.Text = "Prodotti in offerta inclusi";
            this.chkOtrVof.UseVisualStyleBackColor = true;
            this.chkOtrVof.Click += new System.EventHandler(this.chkOtrVof_Click);
            // 
            // chkOtrAnn
            // 
            this.chkOtrAnn.AutoSize = true;
            this.chkOtrAnn.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.chkOtrAnn.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkOtrAnn.Location = new System.Drawing.Point(182, 116);
            this.chkOtrAnn.Name = "chkOtrAnn";
            this.chkOtrAnn.Size = new System.Drawing.Size(77, 19);
            this.chkOtrAnn.TabIndex = 35;
            this.chkOtrAnn.Text = "Annullato";
            this.chkOtrAnn.UseVisualStyleBackColor = true;
            this.chkOtrAnn.Click += new System.EventHandler(this.chkOtrVof_Click);
            // 
            // dtpOtrDtf
            // 
            this.dtpOtrDtf.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dtpOtrDtf.Location = new System.Drawing.Point(448, 235);
            this.dtpOtrDtf.Name = "dtpOtrDtf";
            this.dtpOtrDtf.Size = new System.Drawing.Size(142, 21);
            this.dtpOtrDtf.TabIndex = 34;
            this.dtpOtrDtf.Visible = false;
            this.dtpOtrDtf.Leave += new System.EventHandler(this.dtpOtrDti_Leave);
            // 
            // dtpOtrDti
            // 
            this.dtpOtrDti.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dtpOtrDti.Location = new System.Drawing.Point(324, 235);
            this.dtpOtrDti.Name = "dtpOtrDti";
            this.dtpOtrDti.Size = new System.Drawing.Size(109, 21);
            this.dtpOtrDti.TabIndex = 33;
            this.dtpOtrDti.Visible = false;
            this.dtpOtrDti.Leave += new System.EventHandler(this.dtpOtrDti_Leave);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.Location = new System.Drawing.Point(445, 210);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(20, 15);
            this.label8.TabIndex = 32;
            this.label8.Text = "al ";
            this.label8.Visible = false;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label7.Location = new System.Drawing.Point(323, 210);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(64, 15);
            this.label7.TabIndex = 31;
            this.label7.Text = "Valida dal ";
            this.label7.Visible = false;
            // 
            // txtOtrMix
            // 
            this.txtOtrMix.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtOtrMix.Location = new System.Drawing.Point(514, 115);
            this.txtOtrMix.MaxLength = 5;
            this.txtOtrMix.Name = "txtOtrMix";
            this.txtOtrMix.Size = new System.Drawing.Size(43, 21);
            this.txtOtrMix.TabIndex = 28;
            this.txtOtrMix.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtOtrVal_KeyDown);
            this.txtOtrMix.Validated += new System.EventHandler(this.txtOtrDes_Validated);
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label14.Location = new System.Drawing.Point(445, 120);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(68, 15);
            this.label14.TabIndex = 27;
            this.label14.Text = "Codice mix";
            // 
            // txtOtrSmx
            // 
            this.txtOtrSmx.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtOtrSmx.Location = new System.Drawing.Point(389, 76);
            this.txtOtrSmx.MaxLength = 5;
            this.txtOtrSmx.Name = "txtOtrSmx";
            this.txtOtrSmx.Size = new System.Drawing.Size(43, 21);
            this.txtOtrSmx.TabIndex = 26;
            this.txtOtrSmx.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtOtrSmx.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtOtrVal_KeyDown);
            this.txtOtrSmx.Validated += new System.EventHandler(this.txtOtrDes_Validated);
            // 
            // label13
            // 
            this.label13.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label13.Location = new System.Drawing.Point(333, 74);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(54, 36);
            this.label13.TabIndex = 25;
            this.label13.Text = "Soglia max";
            // 
            // txtOtrSmi
            // 
            this.txtOtrSmi.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtOtrSmi.Location = new System.Drawing.Point(553, 77);
            this.txtOtrSmi.MaxLength = 5;
            this.txtOtrSmi.Name = "txtOtrSmi";
            this.txtOtrSmi.Size = new System.Drawing.Size(43, 21);
            this.txtOtrSmi.TabIndex = 24;
            this.txtOtrSmi.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtOtrSmi.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtOtrVal_KeyDown);
            this.txtOtrSmi.Validated += new System.EventHandler(this.txtOtrDes_Validated);
            // 
            // label12
            // 
            this.label12.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label12.Location = new System.Drawing.Point(463, 75);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(99, 28);
            this.label12.TabIndex = 23;
            this.label12.Text = "Soglia minima punti/sconto";
            // 
            // txtOtrSgl
            // 
            this.txtOtrSgl.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtOtrSgl.Location = new System.Drawing.Point(276, 75);
            this.txtOtrSgl.MaxLength = 5;
            this.txtOtrSgl.Name = "txtOtrSgl";
            this.txtOtrSgl.Size = new System.Drawing.Size(43, 21);
            this.txtOtrSgl.TabIndex = 22;
            this.txtOtrSgl.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtOtrSgl.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtOtrVal_KeyDown);
            this.txtOtrSgl.Validated += new System.EventHandler(this.txtOtrDes_Validated);
            // 
            // label11
            // 
            this.label11.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label11.Location = new System.Drawing.Point(200, 72);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(78, 31);
            this.label11.TabIndex = 21;
            this.label11.Text = "Soglia valore x punti";
            this.label11.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // txtOtrPas
            // 
            this.txtOtrPas.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtOtrPas.Location = new System.Drawing.Point(149, 75);
            this.txtOtrPas.MaxLength = 5;
            this.txtOtrPas.Name = "txtOtrPas";
            this.txtOtrPas.Size = new System.Drawing.Size(43, 21);
            this.txtOtrPas.TabIndex = 20;
            this.txtOtrPas.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtOtrPas.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtOtrVal_KeyDown);
            this.txtOtrPas.Validated += new System.EventHandler(this.txtOtrDes_Validated);
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.Location = new System.Drawing.Point(110, 80);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(41, 15);
            this.label10.TabIndex = 19;
            this.label10.Text = "Passo";
            // 
            // txtOtrVal
            // 
            this.txtOtrVal.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtOtrVal.Location = new System.Drawing.Point(59, 75);
            this.txtOtrVal.MaxLength = 5;
            this.txtOtrVal.Name = "txtOtrVal";
            this.txtOtrVal.Size = new System.Drawing.Size(43, 21);
            this.txtOtrVal.TabIndex = 18;
            this.txtOtrVal.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            this.txtOtrVal.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtOtrVal_KeyDown);
            this.txtOtrVal.Validated += new System.EventHandler(this.txtOtrDes_Validated);
            // 
            // label9
            // 
            this.label9.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label9.Location = new System.Drawing.Point(4, 71);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(48, 33);
            this.label9.TabIndex = 17;
            this.label9.Text = "Valore/punti";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(275, 117);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(31, 15);
            this.label6.TabIndex = 16;
            this.label6.Text = "Tipo";
            // 
            // cmbOtrTip
            // 
            this.cmbOtrTip.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbOtrTip.FormattingEnabled = true;
            this.cmbOtrTip.Location = new System.Drawing.Point(312, 113);
            this.cmbOtrTip.Name = "cmbOtrTip";
            this.cmbOtrTip.Size = new System.Drawing.Size(121, 23);
            this.cmbOtrTip.TabIndex = 15;
            this.cmbOtrTip.SelectionChangeCommitted += new System.EventHandler(this.cmbOtrNeg_SelectionChangeCommitted);
            // 
            // cmbOtrCam
            // 
            this.cmbOtrCam.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbOtrCam.FormattingEnabled = true;
            this.cmbOtrCam.Location = new System.Drawing.Point(475, 43);
            this.cmbOtrCam.Name = "cmbOtrCam";
            this.cmbOtrCam.Size = new System.Drawing.Size(121, 23);
            this.cmbOtrCam.TabIndex = 9;
            this.cmbOtrCam.SelectionChangeCommitted += new System.EventHandler(this.cmbOtrNeg_SelectionChangeCommitted);
            // 
            // cmbOtrNeg
            // 
            this.cmbOtrNeg.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbOtrNeg.FormattingEnabled = true;
            this.cmbOtrNeg.Location = new System.Drawing.Point(59, 41);
            this.cmbOtrNeg.Name = "cmbOtrNeg";
            this.cmbOtrNeg.Size = new System.Drawing.Size(133, 23);
            this.cmbOtrNeg.TabIndex = 8;
            this.cmbOtrNeg.SelectionChangeCommitted += new System.EventHandler(this.cmbOtrNeg_SelectionChangeCommitted);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(4, 47);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(53, 15);
            this.label5.TabIndex = 7;
            this.label5.Text = "Negozio";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(397, 51);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(68, 15);
            this.label4.TabIndex = 6;
            this.label4.Text = "Campagna";
            // 
            // cmbOtrGru
            // 
            this.cmbOtrGru.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbOtrGru.FormattingEnabled = true;
            this.cmbOtrGru.Location = new System.Drawing.Point(262, 43);
            this.cmbOtrGru.Name = "cmbOtrGru";
            this.cmbOtrGru.Size = new System.Drawing.Size(122, 23);
            this.cmbOtrGru.TabIndex = 5;
            this.cmbOtrGru.SelectionChangeCommitted += new System.EventHandler(this.cmbOtrNeg_SelectionChangeCommitted);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(209, 49);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(48, 15);
            this.label3.TabIndex = 4;
            this.label3.Text = "Gruppo";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(120, 16);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(72, 15);
            this.label2.TabIndex = 3;
            this.label2.Text = "Descrizione";
            // 
            // txtOtrDes
            // 
            this.txtOtrDes.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtOtrDes.Location = new System.Drawing.Point(192, 10);
            this.txtOtrDes.MaxLength = 30;
            this.txtOtrDes.Name = "txtOtrDes";
            this.txtOtrDes.Size = new System.Drawing.Size(405, 21);
            this.txtOtrDes.TabIndex = 2;
            this.txtOtrDes.TextChanged += new System.EventHandler(this.txtOtr_TextChanged);
            this.txtOtrDes.Leave += new System.EventHandler(this.txtOtrDes_Leave);
            // 
            // txtOtrCod
            // 
            this.txtOtrCod.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtOtrCod.Location = new System.Drawing.Point(59, 13);
            this.txtOtrCod.MaxLength = 3;
            this.txtOtrCod.Name = "txtOtrCod";
            this.txtOtrCod.ReadOnly = true;
            this.txtOtrCod.Size = new System.Drawing.Size(58, 21);
            this.txtOtrCod.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(5, 16);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(45, 15);
            this.label1.TabIndex = 0;
            this.label1.Text = "Codice";
            // 
            // cmbOtrTca
            // 
            this.cmbOtrTca.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbOtrTca.FormattingEnabled = true;
            this.cmbOtrTca.Location = new System.Drawing.Point(143, 213);
            this.cmbOtrTca.Name = "cmbOtrTca";
            this.cmbOtrTca.Size = new System.Drawing.Size(145, 23);
            this.cmbOtrTca.TabIndex = 45;
            this.cmbOtrTca.SelectionChangeCommitted += new System.EventHandler(this.cmbOtrNeg_SelectionChangeCommitted);
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label18.Location = new System.Drawing.Point(5, 220);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(89, 15);
            this.label18.TabIndex = 46;
            this.label18.Text = "Tipo calcolo su";
            // 
            // frmGesOffTransazione
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(864, 614);
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.btnInv);
            this.Controls.Add(this.btnNew);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.menuStrip1);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.MinimumSize = new System.Drawing.Size(860, 580);
            this.Name = "frmGesOffTransazione";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Offerte su totale scontrino";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmGesOffTransazione_FormClosing);
            this.Load += new System.EventHandler(this.frmGesOffTransazione_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmGesOffTransazione_KeyDown);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv2)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem esciToolStripMenuItem;
        private APOffice.APDataGridView dgv1;
        private System.Windows.Forms.Button btnNew;
        private System.Windows.Forms.Button btnInv;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.TextBox txtOtrCod;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cmbOtrGru;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtOtrDes;
        private System.Windows.Forms.ComboBox cmbOtrCam;
        private System.Windows.Forms.ComboBox cmbOtrNeg;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ComboBox cmbOtrTip;
        private System.Windows.Forms.TextBox txtOtrVal;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox txtOtrPas;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.TextBox txtOtrSgl;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.TextBox txtOtrSmi;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.TextBox txtOtrSmx;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.TextBox txtOtrMix;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Button btnOtrMix;
        private System.Windows.Forms.CheckBox chkOtrVof;
        private System.Windows.Forms.CheckBox chkOtrAnn;
        private System.Windows.Forms.DateTimePicker dtpOtrDtf;
        private System.Windows.Forms.DateTimePicker dtpOtrDti;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox txtOtrRep;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.TextBox txtOtrRen;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.Button btnDayNew;
        private System.Windows.Forms.Label label17;
        private APOffice.APDataGridView dgv2;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.ComboBox cmbOtrTca;
    }
}
