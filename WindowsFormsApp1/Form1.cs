using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using System;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Management;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Timers;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock;
using Button = System.Windows.Forms.Button;
using ComboBox = System.Windows.Forms.ComboBox;
using Excel = Microsoft.Office.Interop.Excel;
using Timer = System.Timers.Timer;

namespace WindowsFormsApp1
{
    public partial class FrontPage : Form
    {
        private SerialPort My_SerialPort;
        private ComboBox ComPortName = new ComboBox();
        private bool Console_receiving = false;
        // Regex 正規表達式 ( Regular Expression )
        //宣告 Regex 忽略大小寫 
        private Regex regex = new Regex(@"Arduino MKRZERO [(](COM\d{1,3})[)]", RegexOptions.IgnoreCase);
        //private bool Counting = false;
        private bool State_Ready = false;
        private bool State_Failing = false;
        private int DataGvRowInd;       //數據網格視圖  行 索引
        private int DataGvColInd;       //數據網格視圖  列 索引
        private bool DataGvLoaded;      //數據網格視圖 載入 狀態
        private bool DataGv_Get_Current_Location = true;
        private bool DataGv_Get_Rows_Location = true;
        private int totalrowCount;
        public int TotalRuns;           //總次數
        public int TotalTimes;          //總時間
        //private bool StartCount;        //開始計數剩餘時間
        private bool StartCount;
        private bool PauseCountdown;
        public Form2 F2 = new Form2();
        private Thread t;
        private string ExcelFilePath = null;   //Excel檔案路徑
        public bool ExcelIsUsing;       //Excel狀態
        private int xlCells_RowInd, xlCells_ColInd;
        public JObject SetJobj;
        delegate void Display(string str);
        delegate void DisplayCount(string str);
        delegate void Control(string cmd);

        public FrontPage()
        {
            InitializeComponent();
            F2.Show();
            F2.MainForm = this;//Form2 to Form1
        }
        private void FrontPage_Load(object sender, EventArgs e)
        {
            //string strPath = System.Environment.CurrentDirectory;
            //string strPath = System.Windows.Forms.Application.StartupPath;//啟動路徑
            string strPath = System.Windows.Forms.Application.ExecutablePath;
            Console.WriteLine(strPath);
            Auto_Connect();
            InitJson();
            //Apply_Settings();
            InitTimer();
            StartCount = true; //測試!
        }
        public void ShowTime(string str)
        {
            label_Time_display.Text = str;
        }
        public void ShowCount(string str)
        {
            textBox_TotalTimes.Text = str;
        }
        public void Command(string cmd)
        {
            textBoxReceive.Text += cmd + "\r\n";
            textBoxReceive.SelectionStart = textBoxReceive.Text.Length;
            textBoxReceive.ScrollToCaret();
            switch (cmd)
            {
                case "S":
                    StartCount = true;
                    label_Time_display.BackColor = Color.FromArgb(0, 225, 255);
                    break;
                case "G":
                    State_Ready = false;
                    label_Time_display.BackColor = Color.FromArgb(128, 255, 128);
                    Write_dataGridView(DataGvRowInd, DataGvColInd, label_Time_display.Text);
                    WriteToExcelWithEpplus(ExcelFilePath, xlCells_RowInd, xlCells_ColInd, label_Time_display.Text);
                    break;
                case "C":
                    label_Time_display.BackColor = Color.Transparent;
                    break;
                case "R":   //Ready
                    State_Ready = true;
                    State_Failing = false;
                    label_Time_display.BackColor = Color.FromArgb(128, 255, 128);
                    label_Time_display.Text = "Ready";
                    break;
                case "F":   //Fail
                    State_Failing = true;
                    State_Ready = false;
                    label_Time_display.BackColor = Color.FromArgb(192, 0, 0);
                    label_Time_display.Text = "XX:XX.XXX";
                    Write_dataGridView(DataGvRowInd, DataGvColInd, "XX:XX.XXX");
                    F2.Round_Time = "XX:XX.XXX";
                    //SaveToExcel(ExcelFilePath, xlCells_RowInd, xlCells_ColInd, label_Time_display.Text);
                    break;
            }
        }
        public void Reset()
        {
            State_Ready = false;
            State_Failing = false;
            label_Time_display.Text = "00:00.000";
            label_Time_display.BackColor = Color.Transparent;
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
                ComPortName.Items.Add(port.GetPropertyValue("DeviceID").ToString());
                // ex: Arduino Uno (COM7)
                PorSelector.Items.Add(port.GetPropertyValue("Caption").ToString());
            }
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
            MatchCollection matches = regex.Matches(PorSelector.Text);
            string portName = "";
            foreach (Match match in matches)
            {
                portName = match.Groups[1].Value;
                break;
            }
            if (portName == "")
            {
                MessageBox.Show("請選則正確的計時器");
                return;
            }

            Console_Connect(portName, 115200);
            if (Console_receiving == true)
            {
                label_ComState.Text = PorSelector.Text + " is Opened,Click again to disconnect";
                label_ComState.BackColor = Color.FromArgb(128, 255, 128);
                //Console.WriteLine(PorSelector.Text + " is Opened,Click again to disconnect");
            }
            else
            {
                label_ComState.Text = "Fail to connect " + PorSelector.Text + ", Check for occupancy and try again";
                label_ComState.BackColor = Color.FromArgb(200, 0, 0);
                //Console.WriteLine("Fail to connect " + PorSelector.Text + ", Check for occupancy and try again");
            }
        }
        private void label_ComState_Click(object sender, EventArgs e)
        {
            if (Console_receiving == true)
            {
                CloseComport();//關閉 Serial Port
                Console_receiving = false;
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
                    if (State_Ready == false) SendString("R\n");
                    break;
                case "button_Command_Fail":
                    if (State_Failing == false) SendString("F\n");
                    break;
                case "button_Command_Restart":
                    Reset();
                    break;
                case "button_Command_3":
                    break;
                case "button_Command_4":
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
                        DataGvRowInd += 1;
                        DataGv_Get_Rows_Location = false;
                        dataGridView1.Focus();
                        dataGridView1.CurrentCell = dataGridView1[DataGvColInd, DataGvRowInd];
                        dataGridView1.BeginEdit(true);
                        DataGv_Get_Rows_Location = true;
                    }
                    break;
                case "button_Round_Previous":
                    if (DataGvLoaded == true)
                    {
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
                        else
                        {
                            ExcelFilePath = string.Empty;
                            label_excel_1.Text = ExcelFilePath;
                        }
                    }
                    break;
                case "button_excel_3":
                    //WriteToExcelWithEpplus(ExcelFilePath, 0, 0, "0");
                    //SaveToExcel(ExcelFilePath, xlCells_RowInd, xlCells_ColInd, textBoxSend.Text);
                    WriteToExcelWithEpplus(ExcelFilePath, xlCells_RowInd, xlCells_ColInd, textBoxSend.Text);
                    Write_dataGridView(DataGvRowInd, DataGvColInd, textBoxSend.Text);
                    break;
                case "button_excel_4":
                    Excel.Application oXL;
                    Excel._Workbook oWB;
                    string path = ExcelFilePath;
                    //Start Excel and get Application object.
                    oXL = new Excel.Application();
                    //Get a new workbook.
                    oWB = oXL.Workbooks.Open(path);
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
                            DataTable dt = null;
                            var pakge = new ExcelPackage(ExcelFilePath);
                            ExcelWorkbook workbook = pakge.Workbook;
                            if (workbook != null)
                            {
                                ExcelWorksheet worksheet = workbook.Worksheets.First();
                                dt = WorksheetToTable(worksheet);
                                dataGridView1.DataSource = dt;
                                totalrowCount = dataGridView1.RowCount - 1;
                                foreach (DataGridViewColumn column in dataGridView1.Columns)
                                { column.SortMode = DataGridViewColumnSortMode.NotSortable; }
                                DataGvLoaded = true;
                            }
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
                    break;
            }
        }
        /////////////////////////////////////////  Button   ///////////////////////////////////////////////
        ///
        /////////////////////////////////////////  ComPort   //////////////////////////////////////////////
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
        private void DoReceive()
        {
            Byte[] buffer = new Byte[1024];
            try
            {

                while (Console_receiving)
                {
                    if (My_SerialPort.BytesToRead > 0)
                    {
                        Stopwatch stopWatch = new Stopwatch();
                        //stopWatch.Start();
                        Int32 length = My_SerialPort.Read(buffer, 0, buffer.Length);
                        Array.Resize(ref buffer, length);

                        string buf = Encoding.UTF8.GetString(buffer);
                        string[] StrArr = buf.Split(new string[] { "\r\n" }, StringSplitOptions.None);
                        Console.WriteLine(buf);
                        bool EnDisplay = true;
                        //string last = StrArr.Last();

                        Display d = new Display(ShowTime);
                        Control c = new Control(Command);
                        foreach (string str in StrArr)
                        {
                            if (str.Length == 8 && EnDisplay)
                            {
                                EnDisplay = false;
                                //Display d = new Display(ShowTime);
                                string time = Format_MilliSecond(HexToUint(str));
                                this.Invoke(d, new Object[] { time });
                            }
                            else if (str.Length == 1)
                            {
                                //Control c = new Control(Command);
                                this.Invoke(c, new Object[] { str });
                            }
                        }
                        Array.Resize(ref buffer, 1024);
                        //stopWatch.Stop();
                        //Console.WriteLine(stopWatch.ElapsedMilliseconds);
                    }
                    Thread.Sleep(20);
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

        private void Auto_Connect()
        {
            string portMsg = "";
            var searcher = new ManagementObjectSearcher("SELECT DeviceID,Caption FROM WIN32_SerialPort");
            foreach (ManagementObject port in searcher.Get().Cast<ManagementObject>())
            {
                portMsg += port.GetPropertyValue("Caption").ToString() + "\n";
            }
            if (Console_receiving == true)
            {
                Console_receiving = false;
                CloseComport();//關閉 Serial Port
            }
            MatchCollection matches = regex.Matches(portMsg);
            string portName = "";
            foreach (Match match in matches)
            {
                portName = match.Groups[1].Value;
                portMsg = match.Groups[0].Value;
                break;
            }
            if (portName == "")
                return;

            Console_Connect(portName, 115200);
            if (Console_receiving == true)
            {
                label_ComState.Text = portMsg + " is Opened,Click again to disconnect";
                label_ComState.BackColor = Color.FromArgb(128, 255, 128);
                //Console.WriteLine(PorSelector.Text + " is Opened,Click again to disconnect");
            }
            else
            {
                label_ComState.Text = "Fail to connect " + portMsg + ", Check for occupancy and try again";
                label_ComState.BackColor = Color.FromArgb(200, 0, 0);
                //Console.WriteLine("Fail to connect " + PorSelector.Text + ", Check for occupancy and try again");
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
        private void WriteToExcelWithEpplus(string Filepath, int Rows, int Columns, string Value)
        {
            if (Filepath != null)
            {
                using (ExcelPackage ep = new ExcelPackage(Filepath))//載入Excel檔案
                {
                    ExcelWorksheet sheet = ep.Workbook.Worksheets[0];//取得Sheet
                    sheet.Cells[Rows, Columns].Value = Value;//寫入文字
                    ep.Save();
                }
            }
        }
        private static DataTable WorksheetToTable(ExcelWorksheet worksheet)
        {
            int rows = worksheet.Dimension.End.Row;
            int cols = worksheet.Dimension.End.Column;
            DataTable dt = new DataTable(worksheet.Name);
            DataRow dr = null;
            DataColumn dc = null;

            //dc = dt.Columns.Add("編號");
            //dc = dt.Columns.Add("值");
            for (int c = 1; c <= cols; c++)
            {
                dc = dt.Columns.Add(worksheet.Cells[1, c].Value.ToString());
                Console.WriteLine(worksheet.Cells[1, c].Value.ToString());
            }
            for (int i = 2; i <= rows; i++)
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
                            dr[j - 1] = worksheet.Cells[i, j].Value.ToString();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message);
                    }
                }
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
                if (DataGv_Get_Rows_Location == true)
                {
                    DataGvRowInd = dataGridView1.CurrentCell.RowIndex;
                }

                if (DataGv_Get_Current_Location == true)
                {
                    DataGvColInd = dataGridView1.CurrentCell.ColumnIndex;
                }
                Console.WriteLine("RowIndex  {0}, ColumnIndex  {1}", DataGvRowInd, DataGvColInd);
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
        //JObject SetJobj = ReadJson("Settings");
        //Console.WriteLine(SetJobj["ClassicMouse"]);
        //Console.WriteLine(SetJobj["ClassicMouse"]["TotalTimes"]);
        //Console.WriteLine(SetJobj["ROBOTRACE"]["TotalTimes"]);
        //SetJobj["ROBOTRACE"]["TotalTimes"] = 50;
        //WriteJson(SetJobj, "Settings");
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
        }
        private void comboBox_SettingPage_class_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Console.WriteLine("目前在json數值為：" + SetJobj[comboBox_SettingPage_class.Text]);
            //Console.WriteLine("TotalTimes：" + SetJobj[comboBox_SettingPage_class.Text]["TotalTimes"]);
            //Console.WriteLine("TotalRound：" + SetJobj[comboBox_SettingPage_class.Text]["TotalRound"]);
            textBox_SettingPage_TotalTimes.Text = (String)SetJobj[comboBox_SettingPage_class.Text]["TotalTimes"];
            comboBox_SettingPage_TotalRound.SelectedItem = (String)SetJobj[comboBox_SettingPage_class.Text]["TotalRound"];
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
            TotalRuns = Int32.Parse(comboBox_SettingPage_TotalRound.Text);
            TotalTimes = Int32.Parse(textBox_SettingPage_TotalTimes.Text);
            Console.WriteLine("Settings：{0} \nTotalRuns: {1}\nTotalTimes{2}", SetJobj["Mode"], TotalRuns, TotalTimes);
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
            if (StartCount && !PauseCountdown)
            {
                TotalTimes -= 1;
                string str = (TotalTimes / 60).ToString() +":"+ (TotalTimes % 60).ToString();
                this.Invoke(d, new object[] { str });
                Console.WriteLine(str);
            }
            //int intHour = e.SignalTime.Hour;
            //int intMinute = e.SignalTime.Minute;
            //int intSecond = e.SignalTime.Second;
            //Console.WriteLine(intHour + ":" + intMinute + "." + intSecond);
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
        /////////////////////////////////////////  Form 2   ///////////////////////////////////////////////
    }
}
