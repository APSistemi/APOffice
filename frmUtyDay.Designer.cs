namespace APOffice
{
    partial class frmUtyDay
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
            this.mpK_Calendar1 = new MPK_Calendar.MPK_Calendar();
            this.btnAbb = new System.Windows.Forms.Button();
            this.btnOk = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // mpK_Calendar1
            // 
            this.mpK_Calendar1.AbbreviateWeekDayHeader = true;
            this.mpK_Calendar1.ActiveMonthColor = System.Drawing.Color.White;
            this.mpK_Calendar1.ApptFont = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Bold);
            this.mpK_Calendar1.BackColor = System.Drawing.Color.LightSteelBlue;
            this.mpK_Calendar1.BackgroundColor = System.Drawing.Color.LightSteelBlue;
            this.mpK_Calendar1.BoldedDateFontColor = System.Drawing.Color.Red;
            this.mpK_Calendar1.BoldedDates = null;
            this.mpK_Calendar1.DisplayWeekendsDarker = true;
            this.mpK_Calendar1.GridColor = System.Drawing.Color.Black;
            this.mpK_Calendar1.HeaderColor = System.Drawing.Color.LightSteelBlue;
            this.mpK_Calendar1.HeaderFont = new System.Drawing.Font("Microsoft Sans Serif", 9F);
            this.mpK_Calendar1.InactiveMonthColor = System.Drawing.Color.Silver;
            this.mpK_Calendar1.intDay = 24;
            this.mpK_Calendar1.intMonth = 5;
            this.mpK_Calendar1.intYear = 2005;
            this.mpK_Calendar1.Location = new System.Drawing.Point(-2, -3);
            this.mpK_Calendar1.Name = "mpK_Calendar1";
            this.mpK_Calendar1.NoApptFont = new System.Drawing.Font("Microsoft Sans Serif", 10F);
            this.mpK_Calendar1.NonselectedDayFontColor = System.Drawing.Color.Black;
            this.mpK_Calendar1.SelectedDate = new System.DateTime(2005, 5, 24, 13, 23, 30, 402);
            this.mpK_Calendar1.SelectedDayColor = System.Drawing.Color.LightSteelBlue;
            this.mpK_Calendar1.SelectedDayFontColor = System.Drawing.Color.White;
            this.mpK_Calendar1.ShowCurrentMonthInDay = false;
            this.mpK_Calendar1.ShowGrid = true;
            this.mpK_Calendar1.ShowPrevNextButton = true;
            this.mpK_Calendar1.Size = new System.Drawing.Size(504, 416);
            this.mpK_Calendar1.TabIndex = 13;
            this.mpK_Calendar1.SelectedDateChanged += new MPK_Calendar.SelectedDateChangedEventHandler(this.mpK_Calendar1_SelectedDateChanged);
            // 
            // btnAbb
            // 
            this.btnAbb.Font = new System.Drawing.Font("Microsoft Sans Serif", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAbb.Location = new System.Drawing.Point(0, 414);
            this.btnAbb.Name = "btnAbb";
            this.btnAbb.Size = new System.Drawing.Size(250, 52);
            this.btnAbb.TabIndex = 14;
            this.btnAbb.Text = "Abbandona";
            this.btnAbb.UseVisualStyleBackColor = true;
            this.btnAbb.Click += new System.EventHandler(this.btnAbb_Click);
            // 
            // btnOk
            // 
            this.btnOk.Font = new System.Drawing.Font("Microsoft Sans Serif", 24F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOk.Location = new System.Drawing.Point(249, 414);
            this.btnOk.Name = "btnOk";
            this.btnOk.Size = new System.Drawing.Size(250, 52);
            this.btnOk.TabIndex = 15;
            this.btnOk.Text = "Conferma";
            this.btnOk.UseVisualStyleBackColor = true;
            this.btnOk.Click += new System.EventHandler(this.btnOk_Click);
            // 
            // frmUtyDay
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(498, 464);
            this.ControlBox = false;
            this.Controls.Add(this.btnOk);
            this.Controls.Add(this.btnAbb);
            this.Controls.Add(this.mpK_Calendar1);
            this.Name = "frmUtyDay";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Load += new System.EventHandler(this.frmUtyDay_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private MPK_Calendar.MPK_Calendar mpK_Calendar1;
        private System.Windows.Forms.Button btnAbb;
        private System.Windows.Forms.Button btnOk;

    }
}