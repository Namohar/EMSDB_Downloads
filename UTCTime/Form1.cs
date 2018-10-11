using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace UTCTime
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //string dt = DateTime.UtcNow.ToString();

            double weeks = (Convert.ToDateTime("10/12/2016") - Convert.ToDateTime("10/31/2016")).TotalDays / 7;
        }
    }
}
