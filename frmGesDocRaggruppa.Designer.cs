namespace APOffice
{
    partial class frmGesDocRaggruppa
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
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.esciToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.utilityToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.fatturaDiRiferimentoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dgv1 = new APOffice.APDataGridView();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.btnOk = new System.Windows.Forms.Button();
            this.btnSeekCli = new System.Windows.Forms.Button();
            this.btnEstrai = new System.Windows.Forms.Button();
            this.btnExl = new System.Windows.Forms.Button();
            this.dtpFatDdo = new System.Windows.Forms.DateTimePicker();
            this.dtpMotFin = new System.Windows.Forms.DateTimePicker();
            this.cmbFatCli = new System.Windows.Forms.ComboBox();
            this.cmbMotCau = new System.Windows.Forms.ComboBox();
            this.cmbFatTpd = new System.Windows.Forms.ComboBox();
            this.lblMotNfa = new System.Windows.Forms.Label();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem,
            this.utilityToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(846, 24);
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
            // utilityToolStripMenuItem
            // 
            this.utilityToolStripMenuItem.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.utilityToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fatturaDiRiferimentoToolStripMenuItem});
            this.utilityToolStripMenuItem.Name = "utilityToolStripMenuItem";
            this.utilityToolStripMenuItem.Size = new System.Drawing.Size(50, 20);
            this.utilityToolStripMenuItem.Text = "Utility";
            // 
            // fatturaDiRiferimentoToolStripMenuItem
            // 
            this.fatturaDiRiferimentoToolStripMenuItem.Name = "fatturaDiRiferimentoToolStripMenuItem";
            this.fatturaDiRiferimentoToolStripMenuItem.Size = new System.Drawing.Size(186, 22);
            this.fatturaDiRiferimentoToolStripMenuItem.Text = "Fattura di riferimento";
            this.fatturaDiRiferimentoToolStripMenuItem.Click += new System.EventHandler(this.fatturaDiRiferimentoToolStripMenuItem_Click);
            // 
            // dgv1
            // 
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.Lavender;
            this.dgv1.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv1.Location = new System.Drawing.Point(12, 211);
            this.dgv1.Name = "dgv1";
            this.dgv1.RowHeadersWidth = 20;
            this.dgv1.Size = new System.Drawing.Size(822, 293);
            this.dgv1.TabIndex = 1;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(12, 131);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(92, 16);
            this.label6.TabIndex = 32;
            this.label6.Text = "Tipo Docum";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(12, 45);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(56, 16);
            this.label5.TabIndex = 31;
            this.label5.Text = "Cliente";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(290, 65);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(55, 16);
            this.label4.TabIndex = 30;
            this.label4.Text = "Fino al";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(12, 89);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(98, 16);
            this.label3.TabIndex = 29;
            this.label3.Text = "Causale Mov";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(471, 132);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(93, 16);
            this.label2.TabIndex = 28;
            this.label2.Text = "Data Fattura";
            // 
            // btnOk
            // 
            this.btnOk.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOk.Location = new System.Drawing.Point(290, 157);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(544, 39);
            this.btnOk.TabIndex = 26;
            this.btnOk.Text = "Generazione Fatture";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // btnSeekCli
            // 
            this.btnSeekCli.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSeekCli.Location = new System.Drawing.Point(570, 31);
            this.btnSeekCli.Name = "btnSeekCli";
            this.btnSeekCli.Size = new System.Drawing.Size(264, 41);
            this.btnSeekCli.TabIndex = 25;
            this.btnSeekCli.Text = "Ricerca Cliente";
            this.btnSeekCli.UseVisualStyleBackColor = true;
            this.btnSeekCli.Click += new System.EventHandler(this.btnSeekCli_Click);
            // 
            // btnEstrai
            // 
            this.btnEstrai.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnEstrai.Location = new System.Drawing.Point(570, 80);
            this.btnEstrai.Name = "btnEstrai";
            this.btnEstrai.Size = new System.Drawing.Size(264, 31);
            this.btnEstrai.TabIndex = 24;
            this.btnEstrai.Text = "Estrazione documenti da fatturare";
            this.btnEstrai.UseVisualStyleBackColor = true;
            this.btnEstrai.Click += new System.EventHandler(this.btnEstrai_Click);
            // 
            // btnExl
            // 
            this.btnExl.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExl.Location = new System.Drawing.Point(12, 157);
            this.btnExl.Name = "btnExl";
            this.btnExl.Size = new System.Drawing.Size(272, 39);
            this.btnExl.TabIndex = 23;
            this.btnExl.Text = "De/seleziona Tutti";
            this.btnExl.UseVisualStyleBackColor = true;
            this.btnExl.Click += new System.EventHandler(this.btnExl_Click);
            // 
            // dtpFatDdo
            // 
            this.dtpFatDdo.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dtpFatDdo.Location = new System.Drawing.Point(570, 122);
            this.dtpFatDdo.Name = "dtpFatDdo";
            this.dtpFatDdo.Size = new System.Drawing.Size(264, 29);
            this.dtpFatDdo.TabIndex = 22;
            // 
            // dtpMotFin
            // 
            this.dtpMotFin.CalendarFont = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dtpMotFin.Font = new System.Drawing.Font("Microsoft Sans Serif", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.dtpMotFin.Location = new System.Drawing.Point(290, 82);
            this.dtpMotFin.Name = "dtpMotFin";
            this.dtpMotFin.Size = new System.Drawing.Size(274, 29);
            this.dtpMotFin.TabIndex = 21;
            // 
            // cmbFatCli
            // 
            this.cmbFatCli.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbFatCli.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbFatCli.FormattingEnabled = true;
            this.cmbFatCli.Location = new System.Drawing.Point(104, 33);
            this.cmbFatCli.Name = "cmbFatCli";
            this.cmbFatCli.Size = new System.Drawing.Size(460, 28);
            this.cmbFatCli.TabIndex = 20;
            this.cmbFatCli.SelectionChangeCommitted += new System.EventHandler(this.cmbFatCli_SelectionChangeCommitted);
            // 
            // cmbMotCau
            // 
            this.cmbMotCau.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbMotCau.FormattingEnabled = true;
            this.cmbMotCau.Location = new System.Drawing.Point(104, 81);
            this.cmbMotCau.Name = "cmbMotCau";
            this.cmbMotCau.Size = new System.Drawing.Size(180, 28);
            this.cmbMotCau.TabIndex = 19;
            // 
            // cmbFatTpd
            // 
            this.cmbFatTpd.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbFatTpd.FormattingEnabled = true;
            this.cmbFatTpd.Location = new System.Drawing.Point(104, 123);
            this.cmbFatTpd.Name = "cmbFatTpd";
            this.cmbFatTpd.Size = new System.Drawing.Size(180, 28);
            this.cmbFatTpd.TabIndex = 18;
            // 
            // lblMotNfa
            // 
            this.lblMotNfa.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lblMotNfa.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblMotNfa.Location = new System.Drawing.Point(333, 125);
            this.lblMotNfa.Name = "lblMotNfa";
            this.lblMotNfa.Size = new System.Drawing.Size(100, 23);
            this.lblMotNfa.TabIndex = 33;
            this.lblMotNfa.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // frmGesDocRaggruppa
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(846, 516);
            this.Controls.Add(this.lblMotNfa);
            this.Controls.Add(this.dtpMotFin);
            this.Controls.Add(this.cmbMotCau);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.btnSeekCli);
            this.Controls.Add(this.btnEstrai);
            this.Controls.Add(this.btnExl);
            this.Controls.Add(this.dtpFatDdo);
            this.Controls.Add(this.cmbFatCli);
            this.Controls.Add(this.cmbFatTpd);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.menuStrip1);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmGesDocRaggruppa";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Raggruppamento documenti";
            this.Load += new System.EventHandler(this.frmGesDocRaggruppa_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmGesDocRaggruppa_KeyDown);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem esciToolStripMenuItem;
        private APOffice.APDataGridView dgv1;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnOk;
        private System.Windows.Forms.Button btnSeekCli;
        private System.Windows.Forms.Button btnEstrai;
        private System.Windows.Forms.Button btnExl;
        private System.Windows.Forms.DateTimePicker dtpFatDdo;
        private System.Windows.Forms.DateTimePicker dtpMotFin;
        private System.Windows.Forms.ComboBox cmbFatCli;
        private System.Windows.Forms.ComboBox cmbMotCau;
        private System.Windows.Forms.ComboBox cmbFatTpd;
        private System.Windows.Forms.ToolStripMenuItem utilityToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem fatturaDiRiferimentoToolStripMenuItem;
        private System.Windows.Forms.Label lblMotNfa;
    }
}
