namespace APOffice
{
    partial class frmUtyCtrlJournal
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
            this.dgv1 = new APOffice.APDataGridView();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.lblPath = new System.Windows.Forms.Label();
            this.btnJou = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.dtpDay = new System.Windows.Forms.DateTimePicker();
            this.label2 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.lblToj = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.txtPos = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.lblTom = new System.Windows.Forms.Label();
            this.chkIniRiga_1 = new System.Windows.Forms.CheckBox();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgv1)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.BackColor = System.Drawing.SystemColors.ControlLight;
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.esciToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(932, 24);
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
            this.dgv1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgv1.Location = new System.Drawing.Point(12, 73);
            this.dgv1.Name = "dgv1";
            this.dgv1.Size = new System.Drawing.Size(908, 516);
            this.dgv1.TabIndex = 1;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(0, 595);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(932, 10);
            this.progressBar1.TabIndex = 2;
            // 
            // lblPath
            // 
            this.lblPath.AutoSize = true;
            this.lblPath.Location = new System.Drawing.Point(118, 34);
            this.lblPath.Name = "lblPath";
            this.lblPath.Size = new System.Drawing.Size(143, 13);
            this.lblPath.TabIndex = 11;
            this.lblPath.Text = "c:\\1aTmp\\Journal\\journal.txt";
            // 
            // btnJou
            // 
            this.btnJou.Location = new System.Drawing.Point(71, 27);
            this.btnJou.Name = "btnJou";
            this.btnJou.Size = new System.Drawing.Size(42, 23);
            this.btnJou.TabIndex = 10;
            this.btnJou.Text = "...";
            this.btnJou.UseVisualStyleBackColor = true;
            this.btnJou.Click += new System.EventHandler(this.btnJou_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(13, 37);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(57, 13);
            this.label1.TabIndex = 9;
            this.label1.Text = "File journal";
            // 
            // dtpDay
            // 
            this.dtpDay.Location = new System.Drawing.Point(387, 29);
            this.dtpDay.Name = "dtpDay";
            this.dtpDay.Size = new System.Drawing.Size(184, 20);
            this.dtpDay.TabIndex = 12;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(291, 35);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(92, 13);
            this.label2.TabIndex = 13;
            this.label2.Text = "Data di riferimento";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(845, 26);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 14;
            this.button1.Text = "Estrazione";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(676, 28);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(71, 13);
            this.label3.TabIndex = 15;
            this.label3.Text = "Totale journal";
            // 
            // lblToj
            // 
            this.lblToj.AutoSize = true;
            this.lblToj.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblToj.Location = new System.Drawing.Point(793, 27);
            this.lblToj.Name = "lblToj";
            this.lblToj.Size = new System.Drawing.Size(16, 16);
            this.lblToj.TabIndex = 16;
            this.lblToj.Text = "0";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(606, 34);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(29, 13);
            this.label4.TabIndex = 17;
            this.label4.Text = "POS";
            // 
            // txtPos
            // 
            this.txtPos.Location = new System.Drawing.Point(637, 27);
            this.txtPos.MaxLength = 2;
            this.txtPos.Name = "txtPos";
            this.txtPos.Size = new System.Drawing.Size(28, 20);
            this.txtPos.TabIndex = 18;
            this.txtPos.Text = "01";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(676, 52);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(87, 13);
            this.label5.TabIndex = 19;
            this.label5.Text = "Totale statistiche";
            // 
            // lblTom
            // 
            this.lblTom.AutoSize = true;
            this.lblTom.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTom.Location = new System.Drawing.Point(793, 52);
            this.lblTom.Name = "lblTom";
            this.lblTom.Size = new System.Drawing.Size(16, 16);
            this.lblTom.TabIndex = 20;
            this.lblTom.Text = "0";
            // 
            // chkIniRiga_1
            // 
            this.chkIniRiga_1.AutoSize = true;
            this.chkIniRiga_1.Location = new System.Drawing.Point(101, 53);
            this.chkIniRiga_1.Name = "chkIniRiga_1";
            this.chkIniRiga_1.Size = new System.Drawing.Size(161, 17);
            this.chkIniRiga_1.TabIndex = 21;
            this.chkIniRiga_1.Text = "Forza riga iniziale journal  a 1";
            this.chkIniRiga_1.UseVisualStyleBackColor = true;
            // 
            // frmUtyCtrlJournal
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(932, 605);
            this.Controls.Add(this.chkIniRiga_1);
            this.Controls.Add(this.lblTom);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.txtPos);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.lblToj);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.dtpDay);
            this.Controls.Add(this.lblPath);
            this.Controls.Add(this.btnJou);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.dgv1);
            this.Controls.Add(this.menuStrip1);
            this.MainMenuStrip = this.menuStrip1;
            this.Name = "frmUtyCtrlJournal";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "frmUtyCtrlJournal";
            this.Load += new System.EventHandler(this.frmUtyCtrlJournal_Load);
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
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.Label lblPath;
        private System.Windows.Forms.Button btnJou;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DateTimePicker dtpDay;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label lblToj;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtPos;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label lblTom;
        private System.Windows.Forms.CheckBox chkIniRiga_1;
    }
}
