using System;
using System.Drawing;
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
            textBox_Run_Time.Text = MainForm.Time_For_F2;
            textBox_Team_Information.Text = MainForm.Team_Information_For_F2;
            MatchCollection matches = regex.Matches(MainForm.Round_For_F2);
            foreach (Match match in matches)
            {
                GroupCollection groups = match.Groups;
                textBox_Round.Text = groups[1].Value.Trim();// MainForm.Round_For_F2;
            }
            textBox_Total_Time.Text = MainForm.TotalTimes_For_F2;
            textBox_Run.Text = MainForm.RunTime_For_F2;
            textBox_Score.Text = MainForm.Score_For_F2;
            textBox_Best_Score.Text = MainForm.BestScore_For_F2;
            bool timeout = false;
            if (textBox_Total_Time.Text == "Time OUT")
                timeout = true;
            if (textBox_Run_Time.Text == "Ready")
                textBox_Run_Time.ForeColor = Color.FromArgb(160, 216, 179);
            else if (textBox_Run_Time.Text == "FAIL")
                textBox_Run_Time.ForeColor = Color.FromArgb(192, 0, 0);
            else if (MainForm.LapStart_For_F2 && timeout)
                textBox_Run_Time.ForeColor = Color.FromArgb(192, 0, 0);
            else
                textBox_Run_Time.ForeColor = Color.FromArgb(21, 2, 1);

            if (timeout)
                textBox_Total_Time.ForeColor = Color.FromArgb(192, 0, 0);
            else
                textBox_Total_Time.ForeColor = Color.FromArgb(21, 2, 1);
        }
        private void Form2_Load(object sender, EventArgs e)
        {
            SendToBack();
            InitTimer();
            //BringToFront();
        }

        private void InitTimer()
        {
            timer = new System.Timers.Timer(33);
            timer.AutoReset = true;
            timer.Enabled = true;
            timer.Elapsed += new System.Timers.ElapsedEventHandler(Timer_Elapsed);
        }
        private void Timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            Get_Data_To_Display d = new Get_Data_To_Display(Flash);
            try
            {
                this.Invoke(d, new Object[] { });
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public string Round_Time//Form1 to Form2
        {
            get { return textBox_Run_Time.Text; }
            set { textBox_Run_Time.Text = value; }
        }
        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            timer.AutoReset = false;
            timer.Enabled = false;
            MainForm.Close();
        }
        protected override bool ProcessCmdKey(ref System.Windows.Forms.Message msg, System.Windows.Forms.Keys keyData)
        {
            int WM_KEYDOWN = 256;
            int WM_SYSKEYDOWN = 260;
            if (msg.Msg == WM_KEYDOWN | msg.Msg == WM_SYSKEYDOWN)
            {
                switch (keyData)
                {
                    case Keys.F11:
                        if (FormBorderStyle == FormBorderStyle.None)
                        {
                            this.FormBorderStyle = FormBorderStyle.Sizable;     //設定邊框樣式
                            this.WindowState = FormWindowState.Normal;          //Normal窗體
                        }
                        else
                        {
                            this.FormBorderStyle = FormBorderStyle.None;        //設定窗體為無邊框樣式
                            this.WindowState = FormWindowState.Maximized;       //最大化窗體
                        }
                        break;
                    case Keys.F12:
                        MessageBox.Show("F12");
                        break;
                }
            }
            return false;
        }
    }
}
