using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Threading;
using System.Diagnostics;

namespace WindowsFormsApp1
{
    public partial class Form2 : Form
    {
        public FrontPage MainForm;//Form2 to Form1
        public Form2()
        {
            InitializeComponent();
        }

        private void Form2_Load(object sender, EventArgs e)
        {

            SendToBack();
            //BringToFront();
        }
        public string Round_Time//Form1 to Form2
        {
            get { return textBox_Round_Time.Text; }
            set { textBox_Round_Time.Text = value; }
        }
    }
}
