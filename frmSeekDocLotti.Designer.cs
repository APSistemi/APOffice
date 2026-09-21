namespace APOffice
{
    partial class frmSeekDocLotti
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
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.esciToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.stampaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.stampaElencoToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.utilitàToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.cancellaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.invioLOTTIABilanciaToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.dgv1 = new APOffice.APDataGridView();
            this.btnNew = new System.Windows.Forms.Button();
            this.btnMdy = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.txtSeek = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.cmbSta = new System.Windows.Forms.ComboBox();
            this.cmbYea = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.menuStrip1.Font = new System.Drawing.Font("Segoe UI", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem,
            this.stampaToolStripMenuItem,
            this.stampaElencoToolStripMenuItem,
            this.utilitàToolStripMenuItem,
            this.invioLOTTIABilanciaToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(787, 33);
            this.menuStrip1.TabIndex = 0;
            this.menuStrip1.Text = "menuStrip1";
            // 
            // esciToolStripMenuItem
            // 
            this.esciToolStripMenuItem.Name = "esciToolStripMenuItem";
            this.esciToolStripMenuItem.Size = new System.Drawing.Size(56, 29);
            this.esciToolStripMenuItem.Text = "Esci";
            this.esciToolStripMenuItem.Click += new System.EventHandler(this.esciToolStripMenuItem_Click);
            // 
            // stampaToolStripMenuItem
            // 
            this.stampaToolStripMenuItem.Name = "stampaToolStripMenuItem";
            this.stampaToolStripMenuItem.Size = new System.Drawing.Size(178, 29);
            this.stampaToolStripMenuItem.Text = "Stampa controllo";
            this.stampaToolStripMenuItem.Click += new System.EventHandler(this.stampaToolStripMenuItem_Click);
            // 
            // stampaElencoToolStripMenuItem
            // 
            this.stampaElencoToolStripMenuItem.Name = "stampaElencoToolStripMenuItem";
            this.stampaElencoToolStripMenuItem.Size = new System.Drawing.Size(154, 29);
            this.stampaElencoToolStripMenuItem.Text = "Stampa elenco";
            this.stampaElencoToolStripMenuItem.Click += new System.EventHandler(this.stampaElencoToolStripMenuItem_Click);
            // 
            // utilitàToolStripMenuItem
            // 
            this.utilitàToolStripMenuItem.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.utilitàToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.cancellaToolStripMenuItem});
            this.utilitàToolStripMenuItem.Name = "utilitàToolStripMenuItem";
            this.utilitàToolStripMenuItem.Size = new System.Drawing.Size(77, 29);
            this.utilitàToolStripMenuItem.Text = "Utilità";
            // 
            // cancellaToolStripMenuItem
            // 
            this.cancellaToolStripMenuItem.Name = "cancellaToolStripMenuItem";
            this.cancellaToolStripMenuItem.Size = new System.Drawing.Size(220, 30);
            this.cancellaToolStripMenuItem.Text = "Cancella LOTTO";
            this.cancellaToolStripMenuItem.Click += new System.EventHandler(this.cancellaToolStripMenuItem_Click);
            // 
            // invioLOTTIABilanciaToolStripMenuItem
            // 
            this.invioLOTTIABilanciaToolStripMenuItem.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.invioLOTTIABilanciaToolStripMenuItem.Enabled = false;
            this.invioLOTTIABilanciaToolStripMenuItem.Name = "invioLOTTIABilanciaToolStripMenuItem";
            this.invioLOTTIABilanciaToolStripMenuItem.Size = new System.Drawing.Size(213, 29);
            this.invioLOTTIABilanciaToolStripMenuItem.Text = "Invio LOTTI a bilancia";
            this.invioLOTTIABilanciaToolStripMenuItem.Click += new System.EventHandler(this.invioLOTTIABilanciaToolStripMenuItem_Click);
            // 
            // dgv1
            // 
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv1.Location = new System.Drawing.Point(6, 121);
            this.dgv1.Name = "dgv1";
            this.dgv1.Size = new System.Drawing.Size(776, 400);
            this.dgv1.TabIndex = 1;
            this.dgv1.DoubleClick += new System.EventHandler(this.dgv1_DoubleClick);
            // 
            // btnNew
            // 
            this.btnNew.Location = new System.Drawing.Point(497, 38);
            this.btnNew.Name = "btnNew";
            this.btnNew.Size = new System.Drawing.Size(140, 31);
            this.btnNew.TabIndex = 2;
            this.btnNew.Text = "NUOVO";
            this.btnNew.UseVisualStyleBackColor = true;
            this.btnNew.Click += new System.EventHandler(this.btnNew_Click);
            // 
            // btnMdy
            // 
            this.btnMdy.Location = new System.Drawing.Point(642, 38);
            this.btnMdy.Name = "btnMdy";
            this.btnMdy.Size = new System.Drawing.Size(140, 31);
            this.btnMdy.TabIndex = 3;
            this.btnMdy.Text = "SELEZIONA";
            this.btnMdy.UseVisualStyleBackColor = true;
            this.btnMdy.Click += new System.EventHandler(this.btnMdy_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(421, 83);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(70, 20);
            this.label1.TabIndex = 6;
            this.label1.Text = "Ricerca";
            // 
            // txtSeek
            // 
            this.txtSeek.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtSeek.Location = new System.Drawing.Point(497, 77);
            this.txtSeek.Name = "txtSeek";
            this.txtSeek.Size = new System.Drawing.Size(285, 26);
            this.txtSeek.TabIndex = 7;
            this.txtSeek.DoubleClick += new System.EventHandler(this.txtSeek_DoubleClick);
            this.txtSeek.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtSeek_KeyDown);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(8, 80);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 20);
            this.label2.TabIndex = 8;
            this.label2.Text = "Stato";
            // 
            // cmbSta
            // 
            this.cmbSta.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbSta.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbSta.FormattingEnabled = true;
            this.cmbSta.Location = new System.Drawing.Point(88, 72);
            this.cmbSta.Name = "cmbSta";
            this.cmbSta.Size = new System.Drawing.Size(269, 28);
            this.cmbSta.TabIndex = 34;
            this.cmbSta.SelectionChangeCommitted += new System.EventHandler(this.cmbSta_SelectionChangeCommitted);
            // 
            // cmbYea
            // 
            this.cmbYea.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbYea.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.cmbYea.FormattingEnabled = true;
            this.cmbYea.Location = new System.Drawing.Point(88, 37);
            this.cmbYea.Name = "cmbYea";
            this.cmbYea.Size = new System.Drawing.Size(116, 28);
            this.cmbYea.TabIndex = 36;
            this.cmbYea.SelectionChangeCommitted += new System.EventHandler(this.cmbYea_SelectionChangeCommitted);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(8, 42);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(51, 20);
            this.label3.TabIndex = 35;
            this.label3.Text = "Anno";
            // 
            // frmSeekDocLotti
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(787, 526);
            this.ControlBox = false;
            this.Controls.Add(this.cmbYea);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.cmbSta);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtSeek);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnMdy);
            this.Controls.Add(this.btnNew);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.menuStrip1);
            this.KeyPreview = true;
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmSeekDocLotti";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Gestione Lotti per documento";
            this.Load += new System.EventHandler(this.frmGesDocLotti_Load);
            this.KeyDown += new System.Windows.Forms.KeyEventHandler(this.frmGesDocLotti_KeyDown);
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
        private System.Windows.Forms.Button btnNew;
        private System.Windows.Forms.Button btnMdy;
        private System.Windows.Forms.ToolStripMenuItem stampaToolStripMenuItem;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtSeek;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cmbSta;
        private System.Windows.Forms.ComboBox cmbYea;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ToolStripMenuItem stampaElencoToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem invioLOTTIABilanciaToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem utilitàToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem cancellaToolStripMenuItem;
    }
}
