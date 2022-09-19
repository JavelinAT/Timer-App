using System;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form2 : Form
    {
        public FrontPage MainForm;//Form2 to Form1
        delegate void Get_Data_To_Display();
        System.Timers.Timer timer;
        private Regex roundRE = new Regex(@"Round (\d)", RegexOptions.IgnoreCase);
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

        private void Form2_FormClosing_1(object sender, FormClosingEventArgs e)
        {
            MainForm.Close();
        }
    }
}
