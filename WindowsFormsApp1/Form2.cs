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
using System.Timers;

namespace WindowsFormsApp1
{
    public partial class Form2 : Form
    {
        public FrontPage MainForm;//Form2 to Form1
        delegate void Get_Data_To_Display();
        System.Timers.Timer timer;
        public Form2()
        {
            InitializeComponent();
            
        }
        public void Flash()
        {
            textBox_Round_Time.Text = MainForm.Time_For_F2;
            textBox_Team_Information.Text = MainForm.Team_Information_For_F2;
            textBox_Round.Text = MainForm.Round_For_F2;
            textBox_Total_Time.Text = MainForm.TotalTimes_For_F2;
        }
        private void Form2_Load(object sender, EventArgs e)
        {
            SendToBack();
            InitTimer();
            //t = new Thread(Display_Data)
            //{
            //    IsBackground = true
            //};
            //t.Start();

            //BringToFront();
        }

        private void InitTimer()
        {
            timer = new System.Timers.Timer(16);
            timer.AutoReset = true;
            timer.Enabled = true;
            timer.Elapsed += new System.Timers.ElapsedEventHandler(Timer_Elapsed);
        }
        private void Timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            Get_Data_To_Display d = new Get_Data_To_Display(Flash);
            this.Invoke(d, new Object[] {  });
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
