using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace date_picker_1
{
    public partial class DatePicker : Form
    {
        //internal object value;

        public DatePicker()
        {
            InitializeComponent();
        }

        public string userDate { get; set; }

        private void DatePicker_Load(object sender, EventArgs e)
        {
            dateTimePicker1.CustomFormat = "yyyy.MM.dd";
            dateTimePicker1.Format = DateTimePickerFormat.Custom;
            dateTimePicker1.Value = DateTime.Now;
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        public void button1_Click(object sender, EventArgs e)
        {
            userDate = dateTimePicker1.Value.Date.ToString("yyyy.MM.dd");
        }

    }
}
