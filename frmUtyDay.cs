using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace APOffice
{
    public partial class frmUtyDay : Form
    {
        public string _strDay = "";

        private Boolean _bolCho = false;

        public frmUtyDay()
        {
            InitializeComponent();
        }

        private void frmUtyDay_Load(object sender, EventArgs e)
        {
            //DateTime[] dts = new DateTime[4];
            //dts[0] = new DateTime(2005, 6, 2);
            //dts[1] = new DateTime(2005, 5, 21);
            //dts[2] = new DateTime(2005, 4, 21);
            //dts[3] = new DateTime(2005, 7, 21);
            //mpK_Calendar1.BoldedDates = dts;
            mpK_Calendar1.SelectedDate = DateTime.Now;

            _strDay = mpK_Calendar1.SelectedDate.ToShortDateString();
            _bolCho = true;
        }

        //private void button1_Click(object sender, EventArgs e)
        //{
        //    // Create a new DateTimePicker.
        //    DateTimePicker dateTimePicker1 = new DateTimePicker();
        //    Controls.AddRange(new Control[] {dateTimePicker1}); 
        //    dateTimePicker1.CalendarFont = new Font("Courier New", 36.00F, FontStyle.Italic, GraphicsUnit.Point, ((Byte)(0)));
        //}

        private void mpK_Calendar1_SelectedDateChanged(object sender, MPK_Calendar.SelectedDateChangedEventArgs e)
        {
            //this.Text = e.SelectedDate.ToShortDateString();
            //if(_bolCho)
            //{
            //    this.Close();
            //}
                _strDay = e.SelectedDate.ToShortDateString();
        }

        private void btnAbb_Click(object sender, EventArgs e)
        {
            _strDay = ""; // this.Text;
            this.Close();
        }

        private void btnOk_Click(object sender, EventArgs e)
        {
            //_strDay = ""; // this.Text;
            this.Close();
        }
    }
}
