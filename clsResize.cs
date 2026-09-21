using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

public class clsResize
{
    List<System.Drawing.Rectangle> _arr_control_storage = new List<System.Drawing.Rectangle>();
    private bool showRowHeader = false;
    public clsResize(Form _form_)
    {
        form = _form_; //the calling form
        _formSize = _form_.ClientSize; //Save initial form size
        _fontsize = _form_.Font.Size; //Font size
    }

    private float _fontsize  { get; set; }

    private System.Drawing.SizeF _formSize {get;set; }

    private Form form { get; set; }

    public void _get_initial_size() //get initial size//
    {
        var _controls = _get_all_controls(form);//call the enumerator
        foreach (Control control in _controls) //Loop through the controls
        {
            _arr_control_storage.Add(control.Bounds); //saves control bounds/dimension            
            //If you have datagridview
            if (control.GetType() == typeof(DataGridView))
                _dgv_Column_Adjust(((DataGridView)control), showRowHeader, 0);
        }
    }

    public void _resize() //Set the resize
    {
        string s = "";
        double _form_ratio_width = (double)form.ClientSize.Width /(double)_formSize.Width; //ratio could be greater or less than 1
        double _form_ratio_height = (double)form.ClientSize.Height / (double)_formSize.Height; // this one too
        var _controls = _get_all_controls(form); //reenumerate the control collection
        int _pos = -1;//do not change this value unless you know what you are doing
        foreach (Control control in _controls)
        {
            s = control.Name;
            if (s.Length > 5 && s.Substring(0, 3) == "cmb")
                Console.WriteLine("aaaa");

            _pos += 1;      //increment by 1;

            // do some math calc
            System.Drawing.Size _controlSize = new System.Drawing.Size((int)(_arr_control_storage[_pos].Width * _form_ratio_width),
                (int)(_arr_control_storage[_pos].Height * _form_ratio_height)); //use for sizing

            System.Drawing.Point _controlposition = new System.Drawing.Point((int)
            (_arr_control_storage[_pos].X * _form_ratio_width), (int)(_arr_control_storage[_pos].Y * _form_ratio_height));//use for location

            //set bounds

            control.Bounds = new System.Drawing.Rectangle(_controlposition, _controlSize); //Put together

            //Assuming you have a datagridview inside a form()
            //if you want to show the row header, replace the false statement of 
            //showRowHeader on top/public declaration to true;
            if (control.GetType() == typeof(DataGridView))
                _dgv_Column_Adjust(((DataGridView)control), showRowHeader, _form_ratio_height);

            s = "";

            if (_form_ratio_width != 0 && _form_ratio_height != 0)
            {

                if (control.Name.Length > 5)
                    s = control.Name.Substring(0, 3);

                if (s == "cmb" || s == "dtp" || s == "txt")
                {
                    double dW = _form_ratio_width;  // *Convert.ToDouble(1.01);
                    double dH = _form_ratio_height; // *Convert.ToDouble(1.01);

                    //Font AutoSize
                    control.Font = new System.Drawing.Font(form.Font.FontFamily,
                        (float)(((Convert.ToDouble(_fontsize) * (dW / 2))) +
                        ((Convert.ToDouble(_fontsize) * dH))));
                }
                else
                {
                    //Font AutoSize
                    control.Font = new System.Drawing.Font(form.Font.FontFamily,
                        (float)(((Convert.ToDouble(_fontsize) * _form_ratio_width) / 2) +
                        ((Convert.ToDouble(_fontsize) * _form_ratio_height) / 2)));
                }
            }
        }
    }

    private void _dgv_Column_Adjust(DataGridView dgv, bool _showRowHeader, double form_ratio_height) //if you have Datagridview 
    //and want to resize the column base on its dimension.
    {
        //int intRowHeader = 0;
        //const int Hscrollbarwidth = 5;
        //if (_showRowHeader)
        //    intRowHeader = dgv.RowHeadersWidth;
        //else
        //    dgv.RowHeadersVisible = false;

        //for (int i = 0; i < dgv.ColumnCount; i++)
        //{
        //    if (dgv.Dock == DockStyle.Fill) //in case the datagridview is docked
        //        dgv.Columns[i].Width = ((dgv.Width - intRowHeader) / dgv.ColumnCount);
        //    else
        //        dgv.Columns[i].Width = ((dgv.Width - intRowHeader - Hscrollbarwidth) / dgv.ColumnCount);
        //}

        //dgv.RowHeadersVisible = false;

        int iHscrollbarwidth = 17;
        int iDgvLen = dgv.Width - iHscrollbarwidth;
        int iColsLen = 0;

        dgv.RowTemplate.Height = 30;

        //if (form_ratio_height != 0)
        //{
        //    Log("***");

        //    Log(form_ratio_height.ToString());
        //    Log(dgv.RowTemplate.Height.ToString());

        //    dgv.RowTemplate.Height = Convert.ToInt16(dgv.RowTemplate.Height * (form_ratio_height));

        //    Log(dgv.RowTemplate.Height.ToString());
        //}

        for (int i = 0; i < dgv.ColumnCount; i++)
        {
            iColsLen += dgv.Columns[i].Width;
        }

        int iDiff = iDgvLen - iColsLen;

        if (iDiff != 0)
        {

            int iLen = 0;

            for (int i = 0; i < dgv.ColumnCount; i++)
            {
                dgv.Columns[i].Width = dgv.Columns[i].Width * iDgvLen / iColsLen;
                iLen += dgv.Columns[i].Width;
            }

            if (dgv.Width != iLen)
            {
                iDiff = dgv.Width - iLen;

                if (form.Text == "Pagamento")
                {
                    dgv.Columns[0].Width = dgv.Columns[0].Width + iDiff - 5;
                }
            }
        }

        Console.WriteLine("aaaa");
    } 

    private static IEnumerable<Control> _get_all_controls(Control c)
    {
        return c.Controls.Cast<Control>().SelectMany(item =>
            _get_all_controls(item)).Concat(c.Controls.Cast<Control>()).Where(control => 
            control.Name != string.Empty);
    }

    public void Log(string strTxt)
    {
        StreamWriter sw = new StreamWriter("C:\\ApProject\\Log\\Pippo.txt", true);
        string str = DateTime.Now.ToString("dd/MM/yyyy HH.mm") + "; " + strTxt;
        sw.Write(str + "\r\n");
        ((TextWriter)sw).Flush();
        sw.Close();
        sw.Dispose();
    }

}
