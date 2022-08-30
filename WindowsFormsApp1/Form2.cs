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
        private Thread t;
        delegate void Get_Data_To_Display(string str);
        public Form2()
        {
            InitializeComponent();
            t = new Thread(Display_Data)
            {
                IsBackground = true
            };
            t.Start();
        }
        private void Display_Data()
        {
            try
            {
                //MainForm
                Thread.Sleep(40);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
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
        public string Team_Information//Form1 to Form2
        {
            get { return textBox_Team_Information.Text; }
            set { textBox_Team_Information.Text = value; }
        }
        public string Round//Form1 to Form2
        {
            get { return textBox_Round.Text; }
            set { textBox_Round.Text = value; }
        }
        public string Total_time//Form1 to Form2
        {
            get { return textBox_Total_Time.Text; }
            set { textBox_Total_Time.Text = value; }
        }
        public string Score//Form1 to Form2
        {
            get { return textBox_Score.Text; }
            set { textBox_Score.Text = value; }
        }
    }
}
