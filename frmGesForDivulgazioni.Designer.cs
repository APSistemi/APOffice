namespace APOffice
{
    partial class frmGesForDivulgazioni
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
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.esciToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.downloadToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.importaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.elaboraToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.utilityToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.conversioneTabelleToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.attivaArticoliDelDocumentoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.stampaEtichetteToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.stampeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.articoliNuoviToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.differenzeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.stampaDocumentoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dgv1 = new APOffice.APDataGridView();
            this.dgv2 = new APOffice.APDataGridView();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.lblDiv = new System.Windows.Forms.Label();
            this.lblCnt = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.lblCos = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
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
            this.downloadToolStripMenuItem,
            this.importaToolStripMenuItem,
            this.elaboraToolStripMenuItem,
            this.utilityToolStripMenuItem,
            this.stampaEtichetteToolStripMenuItem,
            this.stampeToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(917, 24);
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
            // downloadToolStripMenuItem
            // 
            this.downloadToolStripMenuItem.Name = "downloadToolStripMenuItem";
            this.downloadToolStripMenuItem.Size = new System.Drawing.Size(73, 20);
            this.downloadToolStripMenuItem.Text = "Download";
            this.downloadToolStripMenuItem.Click += new System.EventHandler(this.downloadToolStripMenuItem_Click);
            // 
            // importaToolStripMenuItem
            // 
            this.importaToolStripMenuItem.Name = "importaToolStripMenuItem";
            this.importaToolStripMenuItem.Size = new System.Drawing.Size(52, 20);
            this.importaToolStripMenuItem.Text = "Carica";
            this.importaToolStripMenuItem.Click += new System.EventHandler(this.importaToolStripMenuItem_Click);
            // 
            // elaboraToolStripMenuItem
            // 
            this.elaboraToolStripMenuItem.Name = "elaboraToolStripMenuItem";
            this.elaboraToolStripMenuItem.Size = new System.Drawing.Size(68, 20);
            this.elaboraToolStripMenuItem.Text = "Aggiorna";
            this.elaboraToolStripMenuItem.Click += new System.EventHandler(this.elaboraToolStripMenuItem_Click);
            // 
            // utilityToolStripMenuItem
            // 
            this.utilityToolStripMenuItem.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.utilityToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.conversioneTabelleToolStripMenuItem,
            this.attivaArticoliDelDocumentoToolStripMenuItem});
            this.utilityToolStripMenuItem.Name = "utilityToolStripMenuItem";
            this.utilityToolStripMenuItem.Size = new System.Drawing.Size(50, 20);
            this.utilityToolStripMenuItem.Text = "Utility";
            // 
            // conversioneTabelleToolStripMenuItem
            // 
            this.conversioneTabelleToolStripMenuItem.Name = "conversioneTabelleToolStripMenuItem";
            this.conversioneTabelleToolStripMenuItem.Size = new System.Drawing.Size(228, 22);
            this.conversioneTabelleToolStripMenuItem.Text = "Conversione tabelle";
            this.conversioneTabelleToolStripMenuItem.Click += new System.EventHandler(this.conversioneTabelleToolStripMenuItem_Click);
            // 
            // attivaArticoliDelDocumentoToolStripMenuItem
            // 
            this.attivaArticoliDelDocumentoToolStripMenuItem.Name = "attivaArticoliDelDocumentoToolStripMenuItem";
            this.attivaArticoliDelDocumentoToolStripMenuItem.Size = new System.Drawing.Size(228, 22);
            this.attivaArticoliDelDocumentoToolStripMenuItem.Text = "Attiva articoli del documento";
            this.attivaArticoliDelDocumentoToolStripMenuItem.Click += new System.EventHandler(this.attivaArticoliDelDocumentoToolStripMenuItem_Click);
            // 
            // stampaEtichetteToolStripMenuItem
            // 
            this.stampaEtichetteToolStripMenuItem.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.stampaEtichetteToolStripMenuItem.Name = "stampaEtichetteToolStripMenuItem";
            this.stampaEtichetteToolStripMenuItem.Size = new System.Drawing.Size(133, 20);
            this.stampaEtichetteToolStripMenuItem.Text = "Generazione etichette";
            this.stampaEtichetteToolStripMenuItem.Click += new System.EventHandler(this.stampaEtichetteToolStripMenuItem_Click);
            // 
            // stampeToolStripMenuItem
            // 
            this.stampeToolStripMenuItem.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.stampeToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.articoliNuoviToolStripMenuItem,
            this.differenzeToolStripMenuItem,
            this.stampaDocumentoToolStripMenuItem});
            this.stampeToolStripMenuItem.Name = "stampeToolStripMenuItem";
            this.stampeToolStripMenuItem.Size = new System.Drawing.Size(59, 20);
            this.stampeToolStripMenuItem.Text = "Stampe";
            // 
            // articoliNuoviToolStripMenuItem
            // 
            this.articoliNuoviToolStripMenuItem.Name = "articoliNuoviToolStripMenuItem";
            this.articoliNuoviToolStripMenuItem.Size = new System.Drawing.Size(196, 22);
            this.articoliNuoviToolStripMenuItem.Text = "Articoli nuovi/annullati";
            this.articoliNuoviToolStripMenuItem.Click += new System.EventHandler(this.articoliNuoviToolStripMenuItem_Click);
            // 
            // differenzeToolStripMenuItem
            // 
            this.differenzeToolStripMenuItem.Name = "differenzeToolStripMenuItem";
            this.differenzeToolStripMenuItem.Size = new System.Drawing.Size(196, 22);
            this.differenzeToolStripMenuItem.Text = "Differenze anagrafiche";
            this.differenzeToolStripMenuItem.Click += new System.EventHandler(this.differenzeToolStripMenuItem_Click);
            // 
            // stampaDocumentoToolStripMenuItem
            // 
            this.stampaDocumentoToolStripMenuItem.Name = "stampaDocumentoToolStripMenuItem";
            this.stampaDocumentoToolStripMenuItem.Size = new System.Drawing.Size(196, 22);
            this.stampaDocumentoToolStripMenuItem.Text = "Stampa documento";
            this.stampaDocumentoToolStripMenuItem.Click += new System.EventHandler(this.stampaDocumentoToolStripMenuItem_Click);
            // 
            // dgv1
            // 
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.Lavender;
            this.dgv1.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle1;
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv1.Location = new System.Drawing.Point(4, 28);
            this.dgv1.Name = "dgv1";
            this.dgv1.RowHeadersWidth = 20;
            this.dgv1.Size = new System.Drawing.Size(910, 166);
            this.dgv1.TabIndex = 1;
            this.dgv1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv1_CellClick);
            this.dgv1.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.dgv1_CellFormatting);
            this.dgv1.CurrentCellChanged += new System.EventHandler(this.dgv1_CurrentCellChanged);
            // 
            // dgv2
            // 
            dataGridViewCellStyle2.BackColor = System.Drawing.Color.Lavender;
            this.dgv2.AlternatingRowsDefaultCellStyle = dataGridViewCellStyle2;
            this.dgv2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv2.Location = new System.Drawing.Point(3, 198);
            this.dgv2.Name = "dgv2";
            this.dgv2.RowHeadersWidth = 20;
            this.dgv2.Size = new System.Drawing.Size(910, 307);
            this.dgv2.TabIndex = 2;
            this.dgv2.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv2_CellClick);
            this.dgv2.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv2_CellDoubleClick);
            this.dgv2.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgv2_CellEndEdit);
            this.dgv2.CellFormatting += new System.Windows.Forms.DataGridViewCellFormattingEventHandler(this.dgv2_CellFormatting);
            this.dgv2.CurrentCellChanged += new System.EventHandler(this.dgv2_CurrentCellChanged);
            this.dgv2.EditingControlShowing += new System.Windows.Forms.DataGridViewEditingControlShowingEventHandler(this.dgv2_EditingControlShowing);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(5, 589);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(908, 10);
            this.progressBar1.TabIndex = 3;
            // 
            // lblDiv
            // 
            this.lblDiv.AutoSize = true;
            this.lblDiv.Location = new System.Drawing.Point(4, 570);
            this.lblDiv.Name = "lblDiv";
            this.lblDiv.Size = new System.Drawing.Size(16, 13);
            this.lblDiv.TabIndex = 4;
            this.lblDiv.Text = "...";
            // 
            // lblCnt
            // 
            this.lblCnt.AutoSize = true;
            this.lblCnt.Location = new System.Drawing.Point(857, 572);
            this.lblCnt.Name = "lblCnt";
            this.lblCnt.Size = new System.Drawing.Size(16, 13);
            this.lblCnt.TabIndex = 5;
            this.lblCnt.Text = "...";
            // 
            // panel1
            // 
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.panel1.Controls.Add(this.lblCos);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Location = new System.Drawing.Point(4, 512);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(909, 55);
            this.panel1.TabIndex = 6;
            // 
            // lblCos
            // 
            this.lblCos.AutoSize = true;
            this.lblCos.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblCos.Location = new System.Drawing.Point(9, 29);
            this.lblCos.Name = "lblCos";
            this.lblCos.Size = new System.Drawing.Size(0, 16);
            this.lblCos.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(5, 8);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(66, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Storico Costi";
            // 
            // frmGesForDivulgazioni
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(917, 601);
            this.ControlBox = false;
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.lblCnt);
            this.Controls.Add(this.lblDiv);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.dgv2);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.menuStrip1);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmGesForDivulgazioni";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Variazioni da fornitori";
            this.Load += new System.EventHandler(this.frmGesForDivulgazioni_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmGesForDivulgazioni_KeyDown);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dgv2)).EndInit();
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem esciToolStripMenuItem;
        private APOffice.APDataGridView dgv1;
        private System.Windows.Forms.ToolStripMenuItem importaToolStripMenuItem;
        private APOffice.APDataGridView dgv2;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.ToolStripMenuItem elaboraToolStripMenuItem;
        private System.Windows.Forms.Label lblDiv;
        private System.Windows.Forms.Label lblCnt;
        private System.Windows.Forms.ToolStripMenuItem downloadToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem stampaEtichetteToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem utilityToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem conversioneTabelleToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem attivaArticoliDelDocumentoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem stampeToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem articoliNuoviToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem differenzeToolStripMenuItem;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Label lblCos;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ToolStripMenuItem stampaDocumentoToolStripMenuItem;
    }
}
