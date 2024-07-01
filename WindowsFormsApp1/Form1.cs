using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Management;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using TimerLibrary;
using Button = System.Windows.Forms.Button;
using Excel = Microsoft.Office.Interop.Excel;
using Timer = System.Timers.Timer;

namespace WindowsFormsApp1
{
    public partial class FrontPage : Form
    {
        private SerialPort My_SerialPort;
        private bool Console_receiving = false;
        // Regex 正規表達式 ( Regular Expression )
        //宣告 Regex 忽略大小寫 
        private readonly Regex regexMKRZERO = new Regex(@"Arduino MKRZERO [(](COM\d{1,3})[)]", RegexOptions.IgnoreCase);
        private readonly Regex regex = new Regex(@"[(](COM\d{1,3})[)]", RegexOptions.IgnoreCase);
        private readonly Regex ReScore = new Regex(@"\s*(\d{1,2})\s*:\s*(\d{1,2})\s*.\s*(\d{1,3})", RegexOptions.IgnoreCase);
        //private bool Counting = false;
        private string AppName;
        private bool State_Ready = false;
        private bool State_Failing = false;
        private int DataGvRowInd;       //數據網格視圖  行 索引
        private int DataGvColInd;       //數據網格視圖  列 索引
        private bool DataGvLoaded;      //數據網格視圖 載入 狀態
        private bool DataGv_Get_Current_Location = true;
        private bool DataGv_Get_Rows_Location = true;
        private int totalrowCount;
        public bool ClassicMouseMode = false;
        public int ClassicMousebonus = 1;
        public int TotalRuns;           //總次數
        public int TotalTimes;          //總時間
        private bool TimeOut;           //超時表示位元
        private bool StartCount;        //開始計數剩餘時間
        private bool LapStart = false;          //單圈開始
        private bool PauseCountdown;    //暫停倒數
        public Form2 F2 = new Form2();
        private Thread t;
        private string ExcelFilePath = null;        //Excel檔案路徑
        public bool ExcelIsUsing;                   //Excel狀態
        public bool ExcelIsLoaded = false;          //Excel載入完畢
        private int xlCells_RowInd, xlCells_ColInd; //Excel  行、列指標
        //public static Competition Team_List = new Competition();
        //public static Competition Team_List = new Competition("2022");
        public static Competition Team_List = new Competition();
        public JObject SetJobj;
        delegate void Display(string str);
        delegate void DisplayCount(string str);
        //delegate void Control(string cmd);
        delegate void Control(BoardState cmd);

        public FrontPage()
        {
            InitializeComponent();
            F2.MainForm = this;//form2 to form1
            F2.Show();
        }
        private void FrontPage_Load(object sender, EventArgs e)
        {

            //string strPath = System.Environment.CurrentDirectory;
            //string strPath = System.Windows.Forms.Application.StartupPath;//啟動路徑
            string strPath = System.Windows.Forms.Application.ExecutablePath;
            Console.WriteLine(strPath);
            AppName = this.Text;
            Console.WriteLine(AppName);
            ComPort_Connect(null);
            checkedListBox_Setting.SetItemChecked(0, true);
            InitJson();
            InitTimer();
        }
        public void ShowTime(string str)
        {
            label_Time_display.Text = str;
        }
        public void ShowCount(string str)
        {
            textBox_TotalTimes.Text = str;
            if (TimeOut)
                textBox_TotalTimes.ForeColor = Color.FromArgb(192, 0, 0);
            else textBox_TotalTimes.ForeColor = SystemColors.WindowText;
        }
        public int MazeTimeTP;
        public void Command(BoardState cmd)
        {
            textBoxReceive.Text += cmd + "\r\n";
            textBoxReceive.SelectionStart = textBoxReceive.Text.Length;
            textBoxReceive.ScrollToCaret();
            switch (cmd)
            {
                case BoardState.Start:
                    if (ExcelIsLoaded && !StartCount)
                        StartCount = true;
                    if (ExcelIsLoaded && ClassicMouseMode)
                    {
                        Team_List.Team[DataGvRowInd].MazeTime[DataGvColInd - 4].mSec = MazeTimeTP;
                        string strMazeTime = Team_List.Team[DataGvRowInd].MazeTime[DataGvColInd - 4].ToString();
                        Write_dataGridView(DataGvRowInd, DataGvColInd + (TotalRuns * 2), strMazeTime);
                        WriteToExcelWithEpplus(ExcelFilePath, xlCells_RowInd, xlCells_ColInd + (TotalRuns * 2), strMazeTime);
                        Console.WriteLine("MazeTimeTP:{2}\tCount:{0}\tMazeTime{1}\t", Team_List.Team[DataGvRowInd].MazeTime.Count, Team_List.Team[DataGvRowInd].MazeTime[DataGvColInd - 4].mSec, MazeTimeTP);
                    }
                    LapStart = true;
                    label_Time_display.BackColor = Color.FromArgb(0, 225, 255);
                    break;
                case BoardState.End:
                    State_Ready = false;
                    if (TimeOut)
                        label_Time_display.BackColor = Color.FromArgb(255, 0, 0);
                    else
                        label_Time_display.BackColor = Color.FromArgb(128, 255, 128);
                    if (ExcelIsLoaded)
                    {
                        Write_dataGridView(DataGvRowInd, DataGvColInd, label_Time_display.Text);
                        WriteToExcelWithEpplus(ExcelFilePath, xlCells_RowInd, xlCells_ColInd, label_Time_display.Text);
                        WriteToTeamList(DataGvRowInd, DataGvColInd - 4, label_Time_display.Text);
                        if (ClassicMouseMode)
                            Team_List.Team[DataGvRowInd].Score[DataGvColInd - 4].mSec = Team_List.Team[DataGvRowInd].Time[DataGvColInd - 4].mSec + (int)((float)MazeTimeTP / (float)30) - (3000 * ClassicMousebonus);
                        else
                            Team_List.Team[DataGvRowInd].Score[DataGvColInd - 4].mSec = Team_List.Team[DataGvRowInd].Time[DataGvColInd - 4].mSec;
                        string strScore = Team_List.Team[DataGvRowInd].Score[DataGvColInd - 4].ToString();
                        Write_dataGridView(DataGvRowInd, DataGvColInd + TotalRuns, strScore);
                        WriteToExcelWithEpplus(ExcelFilePath, xlCells_RowInd, xlCells_ColInd + TotalRuns, strScore);
                        if (ClassicMouseMode)
                        {
                            Write_dataGridView(DataGvRowInd, DataGvColInd + (TotalRuns * 3), ClassicMousebonus.ToString());
                            WriteToExcelWithEpplus(ExcelFilePath, xlCells_RowInd, xlCells_ColInd + (TotalRuns * 3), ClassicMousebonus.ToString());
                        }
                        ButtonClick(button_Round_Next, null);
                    }
                    LapStart = false;
                    break;
                case BoardState.READY:   //Ready
                    State_Ready = true;
                    State_Failing = false;
                    label_Time_display.BackColor = Color.FromArgb(128, 255, 128);
                    label_Time_display.Text = "Ready";
                    break;
                case BoardState.FAIL:   //Fail
                    State_Ready = false;
                    State_Failing = true;
                    LapStart = false;
                    label_Time_display.BackColor = Color.FromArgb(192, 0, 0);
                    if (ExcelIsLoaded)
                    {
                        string RunTime = label_Time_display.Text;
                        Write_dataGridView(DataGvRowInd, DataGvColInd, RunTime);
                        WriteToExcelWithEpplus(ExcelFilePath, xlCells_RowInd, xlCells_ColInd, RunTime);
                        WriteToTeamList(DataGvRowInd, DataGvColInd - 4, RunTime);
                        //Team_List.Team[DataGvRowInd].Score[DataGvColInd - 4].mSec = Team_List.Team[DataGvRowInd].Time[DataGvColInd - 4].mSec;
                        //string FAILstr = Team_List.Team[DataGvRowInd].Score[DataGvColInd - 4].ToString();
                        Write_dataGridView(DataGvRowInd, DataGvColInd + TotalRuns, RunTime);
                        WriteToExcelWithEpplus(ExcelFilePath, xlCells_RowInd, xlCells_ColInd + TotalRuns, RunTime);
                        ButtonClick(button_Round_Next, null);
                        if (ClassicMouseMode)
                            ClassicMousebonus = 0;
                    }
                    LapStart = false;
                    break;
            }
        }
        //WriteToTeamList
        public void WriteToTeamList(int TeamIndex, int RunIndex, int msScore)
        {
            Team_List.Team[TeamIndex].Time[RunIndex].mSec = msScore;
            Team_List.Team[TeamIndex].Minimum = msScore;
        }
        public void WriteToTeamList(int TeamIndex, int RunIndex, string Score)
        {
            int IntMillisecond;
            if (Score == "FAIL")
            {
                IntMillisecond = StringScore_To_IntMillisecond("99:59.999");
            }
            else
                IntMillisecond = StringScore_To_IntMillisecond(Score);
            Team_List.Team[TeamIndex].Time[RunIndex].mSec = IntMillisecond;
            Team_List.Team[TeamIndex].Score[RunIndex].mSec = IntMillisecond;
            Team_List.Team[TeamIndex].Minimum = IntMillisecond;
        }
        //WriteToTeamList
        public void Reset()
        {
            StartCount = false;
            TimeOut = false;
            State_Ready = false;
            State_Failing = false;
            PauseCountdown = false;
            label_Time_display.Text = "00:00.000";
            label_Time_display.BackColor = Color.Transparent;
            textBox_TotalTimes.ForeColor = SystemColors.WindowText;
            ClassicMousebonus = 1;
            InitJson();
            if (ExcelIsLoaded)
                LodeExcelToDataGridViewWithEpplus(ExcelFilePath);
            if (!Console_receiving)
                ComPort_Connect(null);
            if (Console_receiving)
                SendCommand(COMMAND.RESET);
        }
        public static string Format_MilliSecond(uint TimeMs)
        {
            string T_ms = Convert.ToString(TimeMs % 1000);
            T_ms = T_ms.PadLeft(3, '0');

            string T_Second = Convert.ToString((TimeMs % 60000) / 1000);
            T_Second = T_Second.PadLeft(2, '0');

            string T_minute = Convert.ToString((int)TimeMs / 60000);
            T_minute = T_minute.PadLeft(2, '0');

            string Time = T_minute + ':' + T_Second + '.' + T_ms;
            return Time;
        }
        public static uint HexToUint(string hexString)
        {
            //運算後的位元組長度:16進位數字字串長/2
            byte[] byteOUT = new byte[4];
            uint data;

            for (int i = 0; i < hexString.Length; i += 2)
            {
                //每2位16進位數字轉換為一個10進位整數
                byteOUT[i / 2] = Convert.ToByte(hexString.Substring(i, 2), 16);
            }
            data = BitConverter.ToUInt32(byteOUT, 0);

            return data;
        }
        private void PorSelector_DropDown(object sender, EventArgs e)
        {
            PorSelector.Items.Clear();
            var searcher = new ManagementObjectSearcher("SELECT DeviceID,Caption FROM WIN32_SerialPort");
            foreach (ManagementObject port in searcher.Get().Cast<ManagementObject>())
            {
                PorSelector.Items.Add(port.GetPropertyValue("Caption").ToString()); // ex: Arduino Uno (COM7)
                //Console.WriteLine("ID   {0},Caption {1}", port.GetPropertyValue("DeviceID"), port.GetPropertyValue("Caption"));
            }

            Console.ReadLine();
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            //關閉 Serial Port
            CloseComport();
        }
        private void PorSelector_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (Console_receiving == true)
            {
                Console_receiving = false;
                CloseComport();//關閉 Serial Port
            }
            ComPort_Connect(PorSelector.Text);
        }
        private void label_ComState_Click(object sender, EventArgs e)
        {
            if (Console_receiving == true)
            {
                CloseComport();//關閉 Serial Port
                Console_receiving = false;
                this.Text = AppName;
                label_ComState.Text = "Click to select Com Port";
                label_ComState.BackColor = Color.Transparent;
            }
            else this.PorSelector.DroppedDown = true;
        }
        /////////////////////////////////////////  Button   ///////////////////////////////////////////////
        private void ButtonClick(object sender, EventArgs e)
        {
            Button btn = (Button)sender;
            Console.WriteLine(btn.Text);
            switch (btn.Name)
            {
                ///////////////////////////////////////////////////////////////
                /////  Command   //////////////////////////////////////////////
                case "button_Command_Ready":
                    //if (State_Ready == false) SendString("R\n");
                    if (State_Ready == false) SendCommand(COMMAND.READY);
                    break;
                case "button_Command_Fail":
                    //if (State_Failing == false) SendString("F\n");
                    if (State_Failing == false && LapStart) SendCommand(COMMAND.FAIL_ENDROUND);
                    break;
                case "button_Command_Restart":
                    Reset();
                    break;
                case "button_Command_LoadExcel":
                    using (OpenFileDialog ofd = new OpenFileDialog())
                    {
                        ofd.Filter = "Excel 活頁簿 (*.xlsx)|*.xlsx|Excel 97-2003 (*.xls)|*.xls|文字檔 (Tab 字元分隔) (*.txt)|*.txt";
                        ofd.Title = "Select Excel file";
                        if (ofd.ShowDialog() == DialogResult.OK)
                        {
                            ExcelFilePath = ofd.FileName;
                            label_excel_1.Text = ExcelFilePath;
                            LodeExcelToDataGridViewWithEpplus(ExcelFilePath);
                        }
                    }
                    break;
                case "button_Command_NoBonus":
                    ClassicMousebonus = 0;
                    break;
                /////  Command   //////////////////////////////////////////////
                ///////////////////////////////////////////////////////////////
                /////  dataGridView   /////////////////////////////////////////
                case "button_Inf_Previous":
                    if (DataGvLoaded == true)
                    {
                        DataGvRowInd += -1;
                        DataGv_Get_Rows_Location = false;
                        dataGridView1.Focus();
                        dataGridView1.CurrentCell = dataGridView1[DataGvColInd, DataGvRowInd];
                        dataGridView1.BeginEdit(true);
                        DataGv_Get_Rows_Location = true;
                    }
                    break;
                case "button_Inf_Next":
                    if (DataGvLoaded == true)
                    {
                        if (LapStart)
                        {
                            MessageBox.Show("Objects cannot be changed during timing");
                            return;
                        }

                        DialogResult Result = MessageBox.Show(
                            "This action resets the parameter\r\n" +
                            "do you want to continue", "Notification", 
                            MessageBoxButtons.YesNo);
                        if (Result == DialogResult.No)
                            return;

                        DataGvRowInd += 1;
                        DataGv_Get_Rows_Location = false;
                        dataGridView1.Focus();
                        DataGvColInd = 4;
                        dataGridView1.CurrentCell = dataGridView1[DataGvColInd, DataGvRowInd];
                        dataGridView1.BeginEdit(true);
                        DataGv_Get_Rows_Location = true;

                        StartCount = false;
                        TimeOut = false;
                        State_Ready = false;
                        State_Failing = false;
                        PauseCountdown = false;
                        label_Time_display.Text = "00:00.000";
                        label_Time_display.BackColor = Color.Transparent;
                        textBox_TotalTimes.ForeColor = SystemColors.WindowText;
                        ClassicMousebonus = 1;
                        InitJson();
                        if (Console_receiving)
                            SendCommand(COMMAND.RESET);
                    }
                    break;
                case "button_Round_Previous":
                    if (DataGvLoaded == true)
                    {
                        if (LapStart)
                        {
                            MessageBox.Show("Objects cannot be changed during timing");
                            return;
                        }
                        DataGvColInd += -1;
                        DataGv_Get_Current_Location = false;
                        dataGridView1.Focus();
                        dataGridView1.CurrentCell = dataGridView1[DataGvColInd, DataGvRowInd];
                        dataGridView1.BeginEdit(true);
                        DataGv_Get_Current_Location = true;
                    }
                    break;
                case "button_Round_Next":
                    if (DataGvLoaded == true)
                    {
                        DataGvColInd += 1;
                        DataGv_Get_Current_Location = false;
                        dataGridView1.Focus();
                        dataGridView1.CurrentCell = dataGridView1[DataGvColInd, DataGvRowInd];
                        dataGridView1.BeginEdit(true);
                        DataGv_Get_Current_Location = true;
                    }
                    break;
                case "button_Total_Times":
                    if (PauseCountdown)
                    {
                        PauseCountdown = false;

                        button_Total_Times.Text = "Pause";
                    }
                    else
                    {
                        PauseCountdown = true;
                        button_Total_Times.Text = "Continue";
                        textBox_TotalTimes.Text += "\nPause";
                    }
                    break;
                /////  dataGridView   /////////////////////////////////////////
                ///////////////////////////////////////////////////////////////
                /////  Excel   ////////////////////////////////////////////////
                case "button_excel_1":
                    Create_Excel_template();
                    break;
                case "button_excel_2":
                    using (OpenFileDialog ofd = new OpenFileDialog())
                    {
                        ofd.Filter = "Excel 活頁簿 (*.xlsx)|*.xlsx|Excel 97-2003 (*.xls)|*.xls|文字檔 (Tab 字元分隔) (*.txt)|*.txt";
                        ofd.Title = "Select Excel file";
                        if (ofd.ShowDialog() == DialogResult.OK)
                        {
                            ExcelFilePath = ofd.FileName;
                            label_excel_1.Text = ExcelFilePath;
                            getExcelFile(ExcelFilePath);
                        }
                    }
                    break;
                case "button_excel_3":
                    //WriteToExcelWithEpplus(ExcelFilePath, 0, 0, "0");
                    //SaveToExcel(ExcelFilePath, xlCells_RowInd, xlCells_ColInd, textBoxSend.Text);
                    WriteToExcelWithEpplus(ExcelFilePath, xlCells_RowInd, xlCells_ColInd, label_Time_display.Text);
                    Write_dataGridView(DataGvRowInd, DataGvColInd, label_Time_display.Text);
                    break;
                case "button_excel_4":
                    Excel.Application oXL;
                    Excel._Workbook oWB;
                    string path = ExcelFilePath;
                    //Start Excel and get Application object.
                    oXL = new Excel.Application();
                    try
                    {
                        oWB = oXL.Workbooks.Open(path);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                    //Get a new workbook.
                    //owb = oxl.workbooks.open(path);
                    oXL.Visible = true;
                    oXL.UserControl = true;
                    break;

                case "button_excel_5":
                    using (OpenFileDialog ofd = new OpenFileDialog())
                    {
                        ofd.Filter = "Excel 活頁簿 (*.xlsx)|*.xlsx|Excel 97-2003 (*.xls)|*.xls|文字檔 (Tab 字元分隔) (*.txt)|*.txt";
                        ofd.Title = "Select Excel file";
                        if (ofd.ShowDialog() == DialogResult.OK)
                        {
                            ExcelFilePath = ofd.FileName;
                            LodeExcelToDataGridViewWithEpplus(ExcelFilePath);
                        }
                    }
                    break;
                /////  Excel   ////////////////////////////////////////////////
                ///////////////////////////////////////////////////////////////
                case "button_Sand":
                    SendString(textBoxSend.Text);
                    textBoxSend.Clear();
                    break;
                case "button_Clear":
                    textBoxReceive.Clear();

                    //textBoxReceive.Clear();
                    //foreach (Object item in checkedListBox_Setting.CheckedItems)
                    //{
                    //    textBoxReceive.Text = textBoxReceive.Text + "\r\n" + item.ToString();
                    //}

                    break;
                case "button_setting":
                    byte Setvalue = new byte();
                    int conut = checkedListBox_Setting.Items.Count;
                    for (int i = 0; i < conut; i++)
                    {
                        bool chelist = checkedListBox_Setting.GetItemChecked(i);
                        if (chelist)
                            //Setvalue |= (byte)(1 << (3 - i));
                            Setvalue |= (byte)(1 << 1);
                        else
                            //Setvalue &= (byte)~(1 << (3 - i));
                            unchecked
                            {
                                Setvalue &= (byte)~(1 << 1);
                            }
                    }
                    Setvalue |= 0x60;
                    SendCommand(Setvalue);
                    break;
            }
        }
        /////////////////////////////////////////  Button   ///////////////////////////////////////////////
        ///
        /////////////////////////////////////////  ComPort   //////////////////////////////////////////////
        public void SendCommand(byte command)
        {
            byte[] Setval = new byte[4] { 0x6F, 0xD9, 0x00, 0x03 }; ;
            Setval[2] = command;
            if (Console_receiving == true)
            {
                try
                {
                    My_SerialPort.Write(Setval, 0, 4);
                }
                catch (Exception ex)
                {
                    CloseComport();
                    MessageBox.Show(ex.Message);
                }
            }
        }
        public void SendCommand(COMMAND command)
        {
            byte[] Setval = new byte[4] { 0x6F, 0xD9, 0x00, 0x03 }; ;
            int cmd = (int)command << 4;
            Setval[2] = (byte)cmd;
            Console.WriteLine(Setval[2].ToString());
            if (Console_receiving == true)
            {
                try
                {
                    My_SerialPort.Write(Setval, 0, 4);
                }
                catch (Exception ex)
                {
                    CloseComport();
                    MessageBox.Show(ex.Message);
                }
            }
            else
            {
                MessageBox.Show("Before starting the operation, please complete the connection settings");
                tabControl1.SelectedTab = tabPage_Setting;
                PorSelector.DroppedDown = true;
            }
        }
        public void Console_Connect(string COM, Int32 baud)
        {
            try
            {
                My_SerialPort = new SerialPort();

                if (My_SerialPort.IsOpen)
                {
                    My_SerialPort.Close();
                }

                //設定 Serial Port 參數
                My_SerialPort.PortName = COM;
                My_SerialPort.BaudRate = baud;
                My_SerialPort.DataBits = 8;
                My_SerialPort.StopBits = StopBits.One;

                if (!My_SerialPort.IsOpen)
                {
                    //開啟 Serial Port
                    My_SerialPort.Open();

                    Console_receiving = true;

                    //開啟執行續做接收動作
                    t = new Thread(DoReceive)
                    {
                        IsBackground = true
                    };
                    t.Start();

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }//連接 Console
        public void CloseComport()
        {
            if (Console_receiving == true)
            {
                try
                {
                    My_SerialPort.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }                         //關閉 Console
        [StructLayout(LayoutKind.Explicit)]
        struct ByteUnion
        {
            [FieldOffset(0)]
            public byte b0;
            [FieldOffset(1)]
            public byte b1;
            [FieldOffset(2)]
            public byte b2;
            [FieldOffset(3)]
            public byte b3;

            [FieldOffset(0)]
            public UInt32 uinTime;
        }
        public enum BoardState
        {
            Stan_by,    //  0 待機
            Timing,     //  1 計時中 (該圈)
            Start,      //  2 開始時間
            End,        //  3 結束時間
            FAIL,
            READY = 10
        }
        public enum COMMAND
        {
            COM_Null = 0,   //  0 待機
            READY,          //  1 就緒
            STARTTime,      //  2 開始時間
            ENDTime,        //  3 結束時間
            FAIL_ENDROUND,  //  4結束計時
            RESET,
        };
        private void DoReceive()
        {
            Byte[] buffer = new Byte[1024];
            Byte[] storebuffer = new Byte[0];
            ByteUnion byteU = new ByteUnion();
            int Board_state = 0;
            Stopwatch stopWatch = new Stopwatch();
            Display d = new Display(ShowTime);
            Control cmd = new Control(Command);
            uint StartTime = 0, EndTime = 0, CurrentTime = 0;
            try
            {
                while (Console_receiving)
                {
                    if (My_SerialPort.BytesToRead > 0)
                    {
                        stopWatch.Start();
                        Int32 length = My_SerialPort.Read(buffer, 0, buffer.Length);
                        Array.Resize(ref buffer, length);
                        int recindex = 0;
                        if (storebuffer.Length > 0)
                        {
                            Byte[] temporary = new Byte[1024];
                            buffer.CopyTo(temporary, storebuffer.Length);
                            Array.Copy(storebuffer, 0, temporary, 0, storebuffer.Length);
                            length += storebuffer.Length;
                            Array.Clear(storebuffer, 0, storebuffer.Length);
                            Array.Resize(ref buffer, length);
                            Array.Resize(ref temporary, length);
                            temporary.CopyTo(buffer, 0);
                        }
                        if (length >= 8)
                        {
                            while (recindex + 8 <= length)
                            {
                                if (buffer[recindex] == 0x90 && buffer[recindex + 1] == 0x26 && buffer[recindex + 7] == 0xFC)
                                {

                                    byteU.b0 = buffer[recindex + 2];
                                    byteU.b1 = buffer[recindex + 3];
                                    byteU.b2 = buffer[recindex + 4];
                                    byteU.b3 = buffer[recindex + 5];
                                    Board_state = (byte)(buffer[recindex + 6] >> 4);
                                    if (Board_state == (int)BoardState.READY) { Invoke(cmd, new object[] { Board_state }); }
                                    else if (Board_state == (int)BoardState.Start) { StartTime = byteU.uinTime; MazeTimeTP = (int)StartTime; Invoke(cmd, new object[] { Board_state }); }
                                    else if (Board_state == (int)BoardState.End) 
                                    {
                                        EndTime = byteU.uinTime;
                                        string strTimes = Format_MilliSecond(EndTime - StartTime);
                                        Invoke(d, new object[] { strTimes }); 
                                        Invoke(cmd, new object[] { Board_state }); 
                                    }
                                    else if (Board_state == (int)BoardState.Timing)
                                    {
                                        CurrentTime = byteU.uinTime;
                                        string strTimes = Format_MilliSecond(CurrentTime - StartTime);
                                        Invoke(d, new object[] { strTimes });
                                    }
                                    else if (Board_state == (int)BoardState.FAIL)
                                    {
                                        //Invoke(d, new object[] { "99:59.999" });
                                        Invoke(d, new object[] { "FAIL" });
                                        Invoke(cmd, new object[] { Board_state });
                                    }
                                    Console.WriteLine("sW:{0}\t\tT:{1}\tBs:{2}", stopWatch.ElapsedMilliseconds, byteU.uinTime, (BoardState)Board_state);

                                    recindex += 8;
                                }
                                else recindex += 1;
                            }
                        }
                        Array.Resize(ref storebuffer, 1024);
                        Array.Copy(buffer, recindex, storebuffer, 0, length - recindex);
                        Array.Resize(ref buffer, 1024);
                        Array.Resize(ref storebuffer, length - recindex);

                        stopWatch.Stop();
                        stopWatch.Reset();

                        Thread.Sleep(10);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }               //Console 接收資料
        public void SendString(string send)
        {
            if (send != null && Console_receiving == true)
            {
                try
                {
                    My_SerialPort.Write(send);
                }
                catch (Exception ex)
                {
                    CloseComport();
                    MessageBox.Show(ex.Message);
                }
            }
            else
            {
                MessageBox.Show("Please connect the device before operation");
                tabControl1.SelectedTab = tabPage_Setting;
                this.PorSelector.DroppedDown = true;
            }
        }//Console 發送資料
        private void ComPort_Connect(string PorCaption)
        {
            string portName = null;
            string portMsg = null;
            if (PorCaption != null)
            {
                MatchCollection matches = regex.Matches(PorCaption);
                foreach (Match match in matches)
                {
                    portName = match.Groups[1].Value;
                    break;
                }
                if (portName == null)
                {
                    MessageBox.Show("請選則正確的計時器");
                    return;
                }
            }
            else
            {
                var searcher = new ManagementObjectSearcher("SELECT DeviceID,Caption FROM WIN32_SerialPort");
                foreach (ManagementObject port in searcher.Get().Cast<ManagementObject>())
                {
                    portMsg += port.GetPropertyValue("Caption").ToString() + "\n";
                }
                if (portMsg != null)
                {
                    MatchCollection matches = regexMKRZERO.Matches(portMsg);
                    foreach (Match match in matches)
                    {
                        portName = match.Groups[1].Value;
                        portMsg = match.Groups[0].Value;
                        break;
                    }
                    if (portName == null)
                        return;
                }
                else return;
            }
            if (portName != null)
                Console_Connect(portName, 115200);
            if (Console_receiving == true)
            {
                if (portMsg != null)
                {
                    label_ComState.Text = portMsg + " is Opened,Click again to disconnect";
                    this.Text += "      Connect To " + portMsg;
                }
                else
                {
                    label_ComState.Text = PorCaption + " is Opened,Click again to disconnect";
                    this.Text += "      Connect To " + PorCaption;
                }
                label_ComState.BackColor = Color.FromArgb(128, 255, 128);
            }
            else
            {
                label_ComState.Text = "Fail to connect " + PorSelector.Text + ", Check for occupancy and try again";
                label_ComState.BackColor = Color.FromArgb(200, 0, 0);
            }
        }
        /////////////////////////////////////////  ComPort   //////////////////////////////////////////////
        ///
        /////////////////////////////////////////  Excel   ////////////////////////////////////////////////
        private void Create_Excel_template()
        {
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;
            Excel.Range oRng;

            try
            {
                //Start Excel and get Application object.
                oXL = new Excel.Application
                {
                    Visible = true
                };

                //Get a new workbook.
                oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
                oSheet = (Excel._Worksheet)oWB.ActiveSheet;
                string[] saNames = new string[] { "Order" , "Name" , "ID" , "隊伍名稱" ,
                    "Runtime1" , "Runtime2" , "Runtime3" , "Runtime4" , "Runtime5" ,
                    "Mazetime2" ,"Mazetime3","Mazetime4","Mazetime5",
                    "Score 1","Score 2","Score 3","Score 4","Score 5",
                    "Bouns2","Bouns3","Bouns4","Bouns5","Best score"};
                oSheet.get_Range("A1", "W1").Value2 = saNames;
                for (int i = 2; i < 7; i++)
                {
                    oSheet.Cells[i, 1] = (i - 1);
                }

                oSheet.Cells[2, 2] = "Team-A";
                oSheet.Cells[2, 3] = "A1U-003";
                oSheet.Cells[2, 4] = "Alfa";

                oSheet.Cells[3, 2] = "CV-9";
                oSheet.Cells[3, 3] = "A1U-002";
                oSheet.Cells[3, 4] = "Essex ";

                oSheet.Cells[4, 2] = "CV-6";
                oSheet.Cells[4, 3] = "A1U-004";
                oSheet.Cells[4, 4] = "Enterprise";

                oSheet.Cells[5, 2] = "BB-61";
                oSheet.Cells[5, 3] = "A1U-001";
                oSheet.Cells[5, 4] = "Iowa ";

                oSheet.Cells[6, 2] = "DD-12";
                oSheet.Cells[6, 3] = "A1U-005";
                oSheet.Cells[6, 4] = "丹陽";
                //Format A1:D1 as bold, vertical alignment = center.
                oSheet.get_Range("A1", "W1").Cells.Interior.Color = System.Drawing.Color.FromArgb(255, 140, 0).ToArgb();
                oSheet.get_Range("A1", "W1").VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;

                //AutoFit columns
                oRng = oSheet.get_Range("A1", "W1");
                oRng.EntireColumn.AutoFit();

                oWB.SaveAs(System.Windows.Forms.Application.StartupPath + "\\Template");
                //Make sure Excel is visible and give the user control
                //of Microsoft Excel's lifetime.
                oXL.Visible = true;
                oXL.UserControl = true;
            }
            catch (Exception theException)
            {
                String errorMessage;
                errorMessage = "Error: ";
                errorMessage = String.Concat(errorMessage, theException.Message);
                errorMessage = String.Concat(errorMessage, " Line: ");
                errorMessage = String.Concat(errorMessage, theException.Source);

                MessageBox.Show(errorMessage, "Error");
            }
        }
        private void SaveToExcel(string path, int xlRows, int xlColumns, string Value)
        {
            if (!ExcelIsUsing)
            {
                try
                {
                    ExcelIsUsing = true;
                    Excel.Application xlApp = new Excel.Application();
                    Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(path);
                    Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                    xlApp.Visible = false;
                    xlWorksheet.Cells[xlRows, xlColumns].NumberFormat = "@";
                    xlWorksheet.Cells[xlRows, xlColumns].Value = Value;
                    xlWorkbook.Save();
                    //cleanup
                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    Marshal.ReleaseComObject(xlWorksheet);
                    //close and release
                    xlWorkbook.Close();
                    Marshal.ReleaseComObject(xlWorkbook);

                    //quit and release
                    xlApp.Quit();
                    Marshal.ReleaseComObject(xlApp);
                    ExcelIsUsing = false;
                }
                catch (Exception theException)
                {
                    String errorMessage;
                    errorMessage = "Error: ";
                    errorMessage = String.Concat(errorMessage, theException.Message);
                    errorMessage = String.Concat(errorMessage, " Line: ");
                    errorMessage = String.Concat(errorMessage, theException.Source);
                    MessageBox.Show(errorMessage, "Error");
                    ExcelIsUsing = false;
                }
            }
            else MessageBox.Show("Excel File Is Using");
        }
        public void getExcelFile(string xlFilePath)
        {
            ExcelIsUsing = true;
            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(xlFilePath);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;
            dataGridView1.DataSource = null;
            dataGridView1.Columns.Clear();
            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            totalrowCount = rowCount - 2;
            //--------------------------------------
            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {
                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                    {
                        string strdata = xlRange.Cells[i, j].Value2.ToString();
                        if (i == 1)
                        {
                            dataGridView1.Columns.Add("Cells", strdata);
                        }
                        else
                            this.dataGridView1.Rows[i - 2].Cells[j - 1].Value = strdata;
                    }
                }
                if (i != rowCount) this.dataGridView1.Rows.Add();
            }
            foreach (DataGridViewColumn column in dataGridView1.Columns)

            { column.SortMode = DataGridViewColumnSortMode.NotSortable; }

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            //rule of thumb for releasing com objects:
            //  never use two dots, all COM objects must be referenced and released individually
            //  ex: [somthing].[something].[something] is bad

            //release com objects to fully kill excel process from running in the background
            Marshal.ReleaseComObject(xlRange);
            Marshal.ReleaseComObject(xlWorksheet);

            //close and release
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);

            //quit and release
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
            ExcelIsUsing = false;
            DataGvLoaded = true;
            DataGv_Get_Rows_Location = false;
            DataGv_Get_Current_Location = false;
            dataGridView1.Focus();
            dataGridView1.CurrentCell = dataGridView1[4, 0];
            dataGridView1.BeginEdit(true);
            DataGv_Get_Rows_Location = true;
            DataGv_Get_Current_Location = true;
        }
        /////////////////////////////////////////  Excel   ////////////////////////////////////////////////
        ///
        /////////////////////////////////////////  EPPlus  ////////////////////////////////////////////////
        private void LodeExcelToDataGridViewWithEpplus(string Filepath)
        {
            //DataTable dt = null;
            var pakge = new ExcelPackage(Filepath);
            ExcelWorkbook workbook = pakge.Workbook;
            if (workbook != null)
            {
                ExcelWorksheet worksheet = workbook.Worksheets.First();
                DataTable dt = WorksheetToDataTable(worksheet);
                dataGridView1.DataSource = dt;
                totalrowCount = dataGridView1.RowCount - 1;
                foreach (DataGridViewColumn column in dataGridView1.Columns)
                { column.SortMode = DataGridViewColumnSortMode.NotSortable; }
                DataGvLoaded = true;
                ExcelIsLoaded = true;
            }
        }
        private void Lode_Excel_To_DataGridView_WithEpplus_AND_Create_Data_Class(string Filepath)
        {
            DataTable dt = null;
            var pakge = new ExcelPackage(Filepath);
            ExcelWorkbook workbook = pakge.Workbook;
            if (workbook != null)
            {
                ExcelWorksheet worksheet = workbook.Worksheets.First();
                dt = WorksheetToDataTable(worksheet);
                dataGridView1.DataSource = dt;
                totalrowCount = dataGridView1.RowCount - 1;
                foreach (DataGridViewColumn column in dataGridView1.Columns)
                { column.SortMode = DataGridViewColumnSortMode.NotSortable; }
                DataGvLoaded = true;
                ExcelIsLoaded = true;
            }
        }
        private void WriteToExcelWithEpplus(string Filepath, int Rows, int Columns, string Value)
        {
            if (Filepath != null && ExcelIsLoaded)
            {
                using (ExcelPackage ep = new ExcelPackage(Filepath))//載入Excel檔案
                {
                    ExcelWorksheet sheet = ep.Workbook.Worksheets[0];//取得Sheet
                    sheet.Cells[Rows, Columns].Value = Value;//寫入文字
                    sheet.Cells.AutoFitColumns();
                    ep.Save();
                }
            }
        }
        private DataTable WorksheetToDataTable(ExcelWorksheet worksheet)
        {
            int rows = worksheet.Dimension.End.Row;
            int cols = worksheet.Dimension.End.Column;
            DataTable dt = new DataTable(worksheet.Name);
            Team_List.Team.Clear();
            for (int i = 1; i < rows; i++)
            {
                Team_List.Team.Add(new TEAM("TEAM" + i.ToString()));
                for (int j = 0; j < TotalRuns; j++)
                {
                    Team_List.Team[i - 1].Time.Add(new TimerData(0));
                    //-------------------------------------------------------------------------------
                    Team_List.Team[i - 1].MazeTime.Add(new TimerData(0));
                    //-------------------------------------------------------------------------------
                    Team_List.Team[i - 1].Score.Add(new TimerData(0));
                    //-------------------------------------------------------------------------------
                }
            }
            DataRow dr = null;
            DataColumn dc = null;

            for (int c = 1; c <= cols; c++) //標題  Columns 列  垂直
                dc = dt.Columns.Add(worksheet.Cells[1, c].Value.ToString());

            int count = 0;
            for (int i = 2; i <= rows; i++)// rows 行  水平
            {
                if (i >= 1)
                {
                    dr = dt.Rows.Add();
                }
                for (int j = 1; j <= cols; j++)
                {
                    try
                    {
                        if (worksheet.Cells[i, j].Value != null)
                        {
                            object Value = worksheet.Cells[i, j].Value;
                            dr[j - 1] = Value.ToString();
                            if (worksheet.Cells[1, j].Value.ToString() == "Name")
                                Team_List.Team[count].Name = worksheet.Cells[i, j].Value.ToString();
                            if (worksheet.Cells[1, j].Value.ToString() == "ID")
                                Team_List.Team[count].ID = worksheet.Cells[i, j].Value.ToString();
                            if (worksheet.Cells[1, j].Value.ToString() == "隊伍名稱")
                                Team_List.Team[count].Organize = worksheet.Cells[i, j].Value.ToString();
                            if (j > 4 && j < 5 + TotalRuns)
                            {
                                int IntMillisecond;
                                if (Value.ToString() == "FAIL")
                                    IntMillisecond = 5999999;
                                else
                                    IntMillisecond = StringScore_To_IntMillisecond(Value.ToString());

                                Team_List.Team[count].Time[j - 5].mSec = IntMillisecond;
                                Team_List.Team[count].Minimum = IntMillisecond;
                            }
                            if (j > 4 + TotalRuns && j < 5 + TotalRuns * 2)
                            {
                                int IntMillisecond;
                                if (Value.ToString() == "FAIL")
                                    IntMillisecond = 5999999;
                                else
                                    IntMillisecond = StringScore_To_IntMillisecond(Value.ToString());
                                Team_List.Team[count].Score[j - (5 + TotalRuns)].mSec = IntMillisecond;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
                count++;
            }
            return dt;
        }

        /////////////////////////////////////////  EPPlus  ////////////////////////////////////////////////
        ///
        /////////////////////////////////////////  DataGridView   /////////////////////////////////////////
        public void Write_dataGridView(int GwRows, int GwColumns, string GwScore)
        {
            if (DataGvLoaded)
                dataGridView1.Rows[GwRows].Cells[GwColumns].Value = GwScore;
        }
        private void dataGridView1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (ExcelIsUsing == false)
            {
                if (DataGv_Get_Rows_Location == true && LapStart == false)
                {
                    DataGvRowInd = dataGridView1.CurrentCell.RowIndex;
                }

                if (DataGv_Get_Current_Location == true)
                {
                    DataGvColInd = dataGridView1.CurrentCell.ColumnIndex;
                }
                //Console.WriteLine("RowIndex  {0}, ColumnIndex  {1}", DataGvRowInd, DataGvColInd);
                if (DataGvColInd > (4 + TotalRuns - 1)) DataGvColInd = (4 + TotalRuns - 1);
                else if (DataGvColInd < 4) DataGvColInd = 4;
                if (DataGvRowInd > totalrowCount) DataGvRowInd = totalrowCount;
                else if (DataGvRowInd < 0) DataGvRowInd = 0;
                xlCells_RowInd = DataGvRowInd + 2;
                xlCells_ColInd = DataGvColInd + 1;
                textBox_Round.Text = "Round\r\n " + (DataGvColInd - 3) + "/" + TotalRuns;
                if (dataGridView1.Rows[DataGvRowInd].Cells[2].Value != null && dataGridView1.Rows[DataGvRowInd].Cells[3].Value != null)
                {
                    textBox_Team_Information.Text = dataGridView1.Rows[DataGvRowInd].Cells[2].Value.ToString() +
                        "\r\n" + dataGridView1.Rows[DataGvRowInd].Cells[3].Value.ToString();
                }
            }
        }
        /////////////////////////////////////////  DataGridView   /////////////////////////////////////////
        ///
        /////////////////////////////////////////  AppSettings   //////////////////////////////////////////
        //private void button_SettingPage_Apply_Click(object sender, EventArgs e)
        //{
        //    // 讀取設定檔
        //    Configuration config = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
        //    // 移除指定的AppSettings
        //    config.AppSettings.Settings.Remove(comboBox_SettingPage_class.Text);
        //    config.AppSettings.Settings.Remove("Mode");

        //    // 新增指定的appSettings
        //    config.AppSettings.Settings.Add("Mode", comboBox_SettingPage_class.Text);
        //    config.AppSettings.Settings.Add(comboBox_SettingPage_class.Text, textBox_SettingPage_TotalTimes.Text);
        //    config.AppSettings.Settings.Add(comboBox_SettingPage_class.Text, comboBox_SettingPage_TotalRound.Text);

        //    // 儲存設定
        //    config.Save(ConfigurationSaveMode.Modified);

        //    //套用設定
        //    Apply_Settings();

        //    MessageBox.Show("Applied");
        //}
        //private void Apply_Settings()
        //{
        //    ConfigurationManager.RefreshSection("appSettings");
        //    comboBox_SettingPage_class.SelectedItem = ConfigurationManager.AppSettings["Mode"];
        //    string[] Settings = ConfigurationManager.AppSettings[comboBox_SettingPage_class.Text].Split('\u002C');

        //    TotalRuns = Int32.Parse(Settings[1]);
        //    TotalTimes = Int32.Parse(Settings[0]);
        //    Console.WriteLine(ConfigurationManager.AppSettings["Mode"] + "\tTotalRuns\t" + TotalRuns + "\tTotalTimes\t" + TotalTimes);
        //}
        //private void comboBox_SettingPage_class_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    ConfigurationManager.RefreshSection("appSettings");
        //    Console.WriteLine("目前在App.Config數值為：" + ConfigurationManager.AppSettings[comboBox_SettingPage_class.Text]);
        //    string[] Settings = ConfigurationManager.AppSettings[comboBox_SettingPage_class.Text].Split('\u002C');
        //    Console.WriteLine(Settings);
        //    textBox_SettingPage_TotalTimes.Text = Settings[0];
        //    comboBox_SettingPage_TotalRound.SelectedItem = Settings[1];
        //}
        /////////////////////////////////////////  AppSettings   //////////////////////////////////////////
        ///
        /////////////////////////////////////////  Settings.json   ////////////////////////////////////////
        private void button_SettingPage_Apply_Click(object sender, EventArgs e)
        {
            SetJobj["Mode"] = comboBox_SettingPage_class.Text;
            SetJobj[comboBox_SettingPage_class.Text]["TotalTimes"] = textBox_SettingPage_TotalTimes.Text;
            SetJobj[comboBox_SettingPage_class.Text]["TotalRound"] = comboBox_SettingPage_TotalRound.Text;
            TotalRuns = Int32.Parse(comboBox_SettingPage_TotalRound.Text);
            TotalTimes = Int32.Parse(textBox_SettingPage_TotalTimes.Text);
            Console.WriteLine("TotalRuns" + TotalRuns + "\tTotalTimes" + TotalTimes);
            MessageBox.Show(
                    "Mode\t" + SetJobj["Mode"] + "\r\n" +
                    "TotalTimes\t" + SetJobj[comboBox_SettingPage_class.Text]["TotalTimes"] + "\r\n" +
                    "TotalRound\t" + SetJobj[comboBox_SettingPage_class.Text]["TotalRound"]);
            WriteJson(SetJobj, "Settings");
            if ((string)comboBox_SettingPage_class.SelectedItem == "Classic_Mouse")
                ClassicMouseMode = true;
            else ClassicMouseMode = false;
        }
        private void comboBox_SettingPage_class_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox_SettingPage_TotalTimes.Text = (String)SetJobj[comboBox_SettingPage_class.Text]["TotalTimes"];
            comboBox_SettingPage_TotalRound.SelectedItem = (String)SetJobj[comboBox_SettingPage_class.Text]["TotalRound"];
            //Console.WriteLine("目前在json數值為：" + SetJobj[comboBox_SettingPage_class.Text]);
            //Console.WriteLine("TotalTimes：" + SetJobj[comboBox_SettingPage_class.Text]["TotalTimes"]);
            //Console.WriteLine("TotalRound：" + SetJobj[comboBox_SettingPage_class.Text]["TotalRound"]);
        }
        public JObject ReadJson(String Filename)
        {
            try
            {
                JObject jobj = JObject.Parse(File.ReadAllText(Filename + ".json"));
                return jobj;
            }
            catch (System.IO.FileNotFoundException)
            {
                var json = new
                {
                    Mode = "ROBOTRACE",

                    Classic_Mouse = new
                    {
                        TotalTimes = 420,
                        TotalRound = 7
                    },
                    ROBOTRACE = new
                    {
                        TotalTimes = 180,
                        TotalRound = 3
                    },
                    Line_Mouse = new
                    {
                        TotalTimes = 180,
                        TotalRound = 3
                    },
                    Half_Mouse = new
                    {
                        TotalTimes = 420,
                        TotalRound = 5
                    },
                };
                File.WriteAllText(Filename + ".json", JsonConvert.SerializeObject(json, Formatting.Indented));
                JObject jobj = JObject.Parse(JsonConvert.SerializeObject(json, Formatting.Indented));
                return jobj;
            }
            catch (Newtonsoft.Json.JsonReaderException ex)
            {
                MessageBox.Show(ex.ToString());
                return null;
            }
        }
        public void WriteJson(JObject jobj, String Filename)
        {
            File.WriteAllText(Filename + ".json", JsonConvert.SerializeObject(jobj, Formatting.Indented));
        }
        public void InitJson()
        {
            SetJobj = ReadJson("Settings");
            comboBox_SettingPage_class.SelectedItem = (string)SetJobj["Mode"];
            if ((string)comboBox_SettingPage_class.SelectedItem == "Classic_Mouse" || (string)comboBox_SettingPage_class.SelectedItem == "Half_Mouse")
                ClassicMouseMode = true;
            else ClassicMouseMode = false;
            TotalRuns = Int32.Parse(comboBox_SettingPage_TotalRound.Text);
            TotalTimes = Int32.Parse(textBox_SettingPage_TotalTimes.Text);
            if (TotalTimes % 60 < 10)
                textBox_TotalTimes.Text = (TotalTimes / 60).ToString() + ":0" + (TotalTimes % 60).ToString();
            else
                textBox_TotalTimes.Text = (TotalTimes / 60).ToString() + ":" + (TotalTimes % 60).ToString();

            Console.WriteLine("Settings：{0} \tTotalRuns: {1}\tTotalTimes: {2}\tMode: {3}", SetJobj["Mode"], TotalRuns, TotalTimes, ClassicMouseMode);
        }
        /////////////////////////////////////////  Settings.json   ////////////////////////////////////////
        ///
        /////////////////////////////////////////  timer   ////////////////////////////////////////////////
        private void InitTimer()
        {
            //設定 定時間隔(毫秒為單位)
            int interval = 1000;
            Timer timer = new System.Timers.Timer(interval);
            //設定執行一次（false）或一直執行(true)
            timer.AutoReset = true;
            //設定是否執行System.Timers.Timer.Elapsed事件
            timer.Enabled = true;
            //綁定Elapsed事件
            timer.Elapsed += new System.Timers.ElapsedEventHandler(Timer_Elapsed);
        }
        private void Timer_Elapsed(object sender, System.Timers.ElapsedEventArgs e)
        {
            DisplayCount d = new DisplayCount(ShowCount);
            if (StartCount && !PauseCountdown && !TimeOut)
            {
                TotalTimes -= 1;
                string str;
                if (TotalTimes == 0)
                {
                    TimeOut = true;
                    str = "00:00";
                }
                else
                {
                    if (TotalTimes % 60 < 10)
                    {
                        str = (TotalTimes / 60).ToString() + ":0" + (TotalTimes % 60).ToString();
                    }
                    else
                        str = (TotalTimes / 60).ToString() + ":" + (TotalTimes % 60).ToString();
                }
                this.Invoke(d, new object[] { str });
                //Console.WriteLine(str);
            }
        }
        /////////////////////////////////////////  timer   ////////////////////////////////////////////////
        ///
        /////////////////////////////////////////  Form 2   ///////////////////////////////////////////////
        public string Team_Information_For_F2
        {
            get { return textBox_Team_Information.Text; }
            set { textBox_Team_Information.Text = value; }
        }
        public string Round_For_F2
        {
            get { return textBox_Round.Text; }
            set { textBox_Round.Text = value; }
        }
        public string Time_For_F2
        {
            get { return label_Time_display.Text; }
            set { label_Time_display.Text = value; }
        }
        public string TotalTimes_For_F2
        {
            get { return textBox_TotalTimes.Text; }
            set { textBox_TotalTimes.Text = value; }
        }
        public string Score_For_F2
        {
            get { return Score_with_TeamClass(); }
        }
        public string RunTime_For_F2
        {
            get { return RunTime_with_TeamClass(); }
        }
        public string BestScore_For_F2
        {
            get { return BestScore(); }
        }
        public bool LapStart_For_F2
        {
            get { return LapStart; }
        }
        /////////////////////////////////////////  Form 2   ///////////////////////////////////////////////
        ///
        /////////////////////////////////////////  Score   ////////////////////////////////////////////////
        public int StringScore_To_IntMillisecond(string str)
        {
            int Millisecond = 0;
            MatchCollection matches = ReScore.Matches(str);
            // 一一取出 MatchCollection 內容
            foreach (Match match in matches)
            {
                // 將 Match 內所有值的集合傳給 GroupCollection groups
                GroupCollection groups = match.Groups;
                Millisecond = Convert.ToInt32(groups[1].Value.Trim()) * 60 * 1000 + Convert.ToInt32(groups[2].Value.Trim()) * 1000 + Convert.ToInt32(groups[3].Value.Trim());
                // 印出 Group 內 word 值
                //Console.WriteLine("{0}  {1}", groups[0].Value.Trim(), Millisecond);
                break;
            }
            return Millisecond;
        }
        public string Score()
        {
            string str = "";
            if (DataGvLoaded)
            {
                for (int i = 4; i <= (4 + TotalRuns - 1); i++)
                {
                    if (dataGridView1.Rows[DataGvRowInd].Cells[i].Value != null && dataGridView1.Rows[DataGvRowInd].Cells[i].Value.ToString() != "")
                    {
                        str += "Round-" + (i - 3) + "\t" + dataGridView1.Rows[DataGvRowInd].Cells[i].Value.ToString() + "\r\n";
                    }
                }
            }
            return str;
        }
        public string Score_with_TeamClass()
        {
            string str = "";
            int round = 1;
            if (DataGvLoaded)
            {
                foreach (var item in Team_List.Team[DataGvRowInd].Score)
                {
                    if (item.mSec != 0 && item.mSec != 5999999)
                        str += "S" + round.ToString() + " - " + item.ToString() + "\r\n";
                    else if (item.mSec == 5999999)
                        str += "S" + round.ToString() + " - " + "FAIL" + "\r\n";
                    else
                        str += "S" + round.ToString() + " - " + "\r\n";

                    round++;
                }
            }
            return str;
        }
        public string RunTime_with_TeamClass()
        {
            string str = "";
            int round = 1;
            if (DataGvLoaded)
            {
                foreach (var item in Team_List.Team[DataGvRowInd].Time)
                {
                    if (item.mSec != 0 && item.mSec != 5999999)
                        str += "R" + round.ToString() + " - " + item.ToString() + "\r\n";
                    else if (item.mSec == 5999999)
                        str += "R" + round.ToString() + " - " + "FAIL" + "\r\n";
                    else
                        str += "R" + round.ToString() + " - " + "\r\n";

                    round++;
                }
            }
            return str;
        }
        public string BestScore()
        {
            Competition BestScore = new Competition();
            string str = "";
            int count = 0;
            if (DataGvLoaded)
            {
                foreach (var item in Team_List.Team)
                {
                    BestScore.Team.Add(item);
                }
                BestScore.Team.Sort();
                foreach (var item in BestScore.Team)
                {
                    if (item.Minimum != 0 && item.Minimum != 5999999)
                    {
                        count++;

                        string Org = item.Organize;
                        string teamOrg = "";
                        if (Org.Length > 7)
                            for (int i = 0; i < 7; i++)
                                teamOrg += Org[i];
                        else teamOrg = Org;

                        str += count.ToString() + "." + teamOrg + "\t" + item.ToString() + "\r\n";
                        if (count == 9) break;
                    }
                }
            }
            return str;
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 0)
            {
                if (!LapStart)
                {
                    InitJson();
                    if (ExcelIsLoaded)
                        LodeExcelToDataGridViewWithEpplus(ExcelFilePath);
                }
                if (DataGvLoaded == true)
                {
                    DataGv_Get_Current_Location = false;
                    dataGridView1.Focus();
                    dataGridView1.CurrentCell = dataGridView1[DataGvColInd, DataGvRowInd];
                    dataGridView1.BeginEdit(true);
                    DataGv_Get_Current_Location = true;
                }
            }
        }

        public class SCORE
        {
            public int Millisecond { get; set; }
            public string Scorestr
            {
                set { Scorestr = value; }
            }
        }
        /////////////////////////////////////////  Score   ////////////////////////////////////////////////
    }
}
