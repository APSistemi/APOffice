namespace APOffice
{
    partial class frmUtyDivApPhone
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
            this.esciToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.invioASedeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dgv1 = new APOffice.APDataGridView();
            this.btnFill = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.dtpDti = new System.Windows.Forms.DateTimePicker();
            this.dtpDtf = new System.Windows.Forms.DateTimePicker();
            this.label3 = new System.Windows.Forms.Label();
            this.cmbTip = new System.Windows.Forms.ComboBox();
            this.chkAll = new System.Windows.Forms.CheckBox();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.btnSave = new System.Windows.Forms.Button();
            this.dgv0 = new APOffice.APDataGridView();
            this.cmbNeg = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.cmbSta = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.btnCanc = new System.Windows.Forms.Button();
            this.lblCnt = new System.Windows.Forms.Label();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv0)).BeginInit();
            this.SuspendLayout();
            // 
            // esciToolStripMenuItem
            // 
            this.esciToolStripMenuItem.Name = "esciToolStripMenuItem";
            this.esciToolStripMenuItem.Size = new System.Drawing.Size(39, 20);
            this.esciToolStripMenuItem.Text = "Esci";
            this.esciToolStripMenuItem.Click += new System.EventHandler(this.esciToolStripMenuItem_Click);
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem,
            this.invioASedeToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(877, 24);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // invioASedeToolStripMenuItem
            // 
            this.invioASedeToolStripMenuItem.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.invioASedeToolStripMenuItem.Name = "invioASedeToolStripMenuItem";
            this.invioASedeToolStripMenuItem.Size = new System.Drawing.Size(82, 20);
            this.invioASedeToolStripMenuItem.Text = "Invio a Sede";
            this.invioASedeToolStripMenuItem.Click += new System.EventHandler(this.invioASedeToolStripMenuItem_Click);
            // 
            // dgv1
            // 
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv1.Location = new System.Drawing.Point(3, 352);
            this.dgv1.Name = "dgv1";
            this.dgv1.RowHeadersWidth = 20;
            this.dgv1.Size = new System.Drawing.Size(870, 279);
            this.dgv1.TabIndex = 1;
            this.dgv1.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv1_CellDoubleClick);
            // 
            // btnFill
            // 
            this.btnFill.Location = new System.Drawing.Point(749, 26);
            this.btnFill.Name = "btnFill";
            this.btnFill.Size = new System.Drawing.Size(124, 30);
            this.btnFill.TabIndex = 2;
            this.btnFill.Text = "Carica";
            this.btnFill.UseVisualStyleBackColor = true;
            this.btnFill.Click += new System.EventHandler(this.btnFill_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 40);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(56, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Data inizio";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(8, 77);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(50, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "Data fine";
            // 
            // dtpDti
            // 
            this.dtpDti.CalendarFont = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dtpDti.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dtpDti.Location = new System.Drawing.Point(70, 31);
            this.dtpDti.Name = "dtpDti";
            this.dtpDti.Size = new System.Drawing.Size(197, 22);
            this.dtpDti.TabIndex = 5;
            // 
            // dtpDtf
            // 
            this.dtpDtf.CalendarFont = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dtpDtf.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dtpDtf.Location = new System.Drawing.Point(70, 68);
            this.dtpDtf.Name = "dtpDtf";
            this.dtpDtf.Size = new System.Drawing.Size(197, 22);
            this.dtpDtf.TabIndex = 6;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(436, 39);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(28, 13);
            this.label3.TabIndex = 7;
            this.label3.Text = "Tipo";
            // 
            // cmbTip
            // 
            this.cmbTip.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbTip.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbTip.FormattingEnabled = true;
            this.cmbTip.Location = new System.Drawing.Point(488, 29);
            this.cmbTip.Name = "cmbTip";
            this.cmbTip.Size = new System.Drawing.Size(143, 23);
            this.cmbTip.TabIndex = 8;
            // 
            // chkAll
            // 
            this.chkAll.AutoSize = true;
            this.chkAll.Location = new System.Drawing.Point(9, 100);
            this.chkAll.Name = "chkAll";
            this.chkAll.Size = new System.Drawing.Size(248, 17);
            this.chkAll.TabIndex = 61;
            this.chkAll.Text = "Cambia stato su tutte le righe (Stato \'S\' = attive)";
            this.chkAll.UseVisualStyleBackColor = true;
            this.chkAll.Click += new System.EventHandler(this.chkAll_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(1, 653);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(875, 10);
            this.progressBar1.TabIndex = 62;
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(749, 57);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(124, 30);
            this.btnSave.TabIndex = 63;
            this.btnSave.Text = "Aggiorna";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // dgv0
            // 
            this.dgv0.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv0.Location = new System.Drawing.Point(4, 123);
            this.dgv0.Name = "dgv0";
            this.dgv0.RowHeadersWidth = 20;
            this.dgv0.Size = new System.Drawing.Size(870, 221);
            this.dgv0.TabIndex = 64;
            this.dgv0.CurrentCellChanged += new System.EventHandler(this.dgv0_CurrentCellChanged);
            this.dgv0.DoubleClick += new System.EventHandler(this.dgv0_DoubleClick);
            // 
            // cmbNeg
            // 
            this.cmbNeg.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbNeg.FormattingEnabled = true;
            this.cmbNeg.Location = new System.Drawing.Point(488, 91);
            this.cmbNeg.Name = "cmbNeg";
            this.cmbNeg.Size = new System.Drawing.Size(143, 21);
            this.cmbNeg.TabIndex = 68;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(436, 99);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(46, 13);
            this.label4.TabIndex = 67;
            this.label4.Text = "Negozio";
            // 
            // cmbSta
            // 
            this.cmbSta.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSta.FormattingEnabled = true;
            this.cmbSta.Location = new System.Drawing.Point(488, 60);
            this.cmbSta.Name = "cmbSta";
            this.cmbSta.Size = new System.Drawing.Size(143, 21);
            this.cmbSta.TabIndex = 66;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(436, 68);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(32, 13);
            this.label5.TabIndex = 65;
            this.label5.Text = "Stato";
            // 
            // btnCanc
            // 
            this.btnCanc.Location = new System.Drawing.Point(749, 89);
            this.btnCanc.Name = "btnCanc";
            this.btnCanc.Size = new System.Drawing.Size(124, 30);
            this.btnCanc.TabIndex = 69;
            this.btnCanc.Text = "Cancella";
            this.btnCanc.UseVisualStyleBackColor = true;
            this.btnCanc.Click += new System.EventHandler(this.btnCanc_Click);
            // 
            // lblCnt
            // 
            this.lblCnt.AutoSize = true;
            this.lblCnt.Location = new System.Drawing.Point(820, 635);
            this.lblCnt.Name = "lblCnt";
            this.lblCnt.Size = new System.Drawing.Size(13, 13);
            this.lblCnt.TabIndex = 70;
            this.lblCnt.Text = "0";
            // 
            // frmUtyDivApPhone
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(877, 665);
            this.ControlBox = false;
            this.Controls.Add(this.lblCnt);
            this.Controls.Add(this.btnCanc);
            this.Controls.Add(this.cmbNeg);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.cmbSta);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.dgv0);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.chkAll);
            this.Controls.Add(this.cmbTip);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.dtpDtf);
            this.Controls.Add(this.dtpDti);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnFill);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.menuStrip1);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmUtyDivApPhone";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Divulgazione APphone";
            this.Load += new System.EventHandler(this.frmUtyDivApPhone_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmUtyDivApPhone_KeyDown);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv0)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ToolStripMenuItem esciToolStripMenuItem;
        private System.Windows.Forms.MenuStrip menuStrip1;
        private APOffice.APDataGridView dgv1;
        private System.Windows.Forms.Button btnFill;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.DateTimePicker dtpDti;
        private System.Windows.Forms.DateTimePicker dtpDtf;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox cmbTip;
        private System.Windows.Forms.CheckBox chkAll;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.ToolStripMenuItem invioASedeToolStripMenuItem;
        private APOffice.APDataGridView dgv0;
        private System.Windows.Forms.ComboBox cmbNeg;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox cmbSta;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button btnCanc;
        private System.Windows.Forms.Label lblCnt;

    }
}
