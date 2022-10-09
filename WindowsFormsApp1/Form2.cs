﻿using System;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace WindowsFormsApp1
{
    public partial class Form2 : Form
    {
        public FrontPage MainForm;//Form2 to Form1
        delegate void Get_Data_To_Display();
        System.Timers.Timer timer;
        private Regex regex = new Regex(@"Round\s*(\d\s*[/]\s*\d)", RegexOptions.IgnoreCase);
        public Form2()
        {
            InitializeComponent();
        }
        public void Flash()
        {
            textBox_Round_Time.Text = MainForm.Time_For_F2;
            textBox_Team_Information.Text = MainForm.Team_Information_For_F2;
            MatchCollection matches = regex.Matches(MainForm.Round_For_F2);
            // 一一取出 MatchCollection 內容
            foreach (Match match in matches)
            {
                // 將 Match 內所有值的集合傳給 GroupCollection groups
                GroupCollection groups = match.Groups;
                // 印出 Group 內 word 值
                Console.WriteLine(groups[1].Value.Trim());
                textBox_Round.Text = groups[1].Value.Trim();// MainForm.Round_For_F2;
            }
            textBox_Total_Time.Text = MainForm.TotalTimes_For_F2;
        }
        private void Form2_Load(object sender, EventArgs e)
        {
            SendToBack();
            InitTimer();
            //BringToFront();
        }

        private void InitTimer()
        {
            timer = new System.Timers.Timer(20);
            timer.AutoReset = true;
            timer.Enabled = true;
            timer.Elapsed += new System.Timers.ElapsedEventHandler(Timer_Elapsed);
        }
        private void Timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            Get_Data_To_Display d = new Get_Data_To_Display(Flash);
            this.Invoke(d, new Object[] { });
        }

        public string Round_Time//Form1 to Form2
        {
            get { return textBox_Round_Time.Text; }
            set { textBox_Round_Time.Text = value; }
        }
        //public string Team_Information//Form1 to Form2
        //{
        //    get { return textBox_Team_Information.Text; }
        //    set { textBox_Team_Information.Text = value; }
        //}
        //public string Round//Form1 to Form2
        //{
        //    get { return textBox_Round.Text; }
        //    set { textBox_Round.Text = value; }
        //}
        //public string Total_time//Form1 to Form2
        //{
        //    get { return textBox_Total_Time.Text; }
        //    set { textBox_Total_Time.Text = value; }
        //}
        //public string Score//Form1 to Form2
        //{
        //    get { return textBox_Score.Text; }
        //    set { textBox_Score.Text = value; }
        //}

        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            timer.AutoReset = false;
            timer.Enabled = false;
            MainForm.Close();
        }
        private void button1_Click(object sender, EventArgs e)
        {
            if (FormBorderStyle == FormBorderStyle.None)
            {
                this.FormBorderStyle = FormBorderStyle.Sizable;     //設定窗體為無邊框樣式
                this.WindowState = FormWindowState.Normal;    //最大化窗體
            }
            else
            {
                this.FormBorderStyle = FormBorderStyle.None;     //設定窗體為無邊框樣式
                this.WindowState = FormWindowState.Maximized;    //最大化窗體
            }
        }
    }
}
