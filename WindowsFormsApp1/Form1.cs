using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Data.OleDb;
using System.Management;
using System.Diagnostics;

namespace WindowsFormsApp1
{
    public partial class FrontPage : Form
    {
        private SerialPort My_SerialPort;
        private ComboBox ComPortName = new ComboBox();
        private bool Console_receiving = false;
        //private bool Counting = false;
        private bool State_Ready = false;
        private bool State_Failing = false;
        private int DataGvRowInd;
        private int DataGvColInd;
        private bool DataGvLoaded;
        private bool DataGv_Get_Current_Location = true;
        private bool DataGv_Get_Rows_Location = true;
        private int totalrowCount;
        public int TotalRuns = 3;
        public Form2 F2 = new Form2();
        private Thread t;
        private string ExcelFilePath;
        public bool ExcelIsUsing;
        private int xlCells_RowInd, xlCells_ColInd;
        delegate void Display(string str);
        delegate void Control(string cmd);

        public FrontPage()
        {
            InitializeComponent();
            F2.Show();

            F2.MainForm = this;//Form2 to Form1
        }
        public void ShowTime(string str)//3
        {
            label_Time_display.Text = str;
            F2.Round_Time = str;
        }
        public void Command(string cmd)
        {
            textBoxReceive.Text += cmd + "\r\n";
            textBoxReceive.SelectionStart = textBoxReceive.Text.Length;
            textBoxReceive.ScrollToCaret();
            switch (cmd)
            {
                case "S":
                    //Counting = true;
                    label_Time_display.BackColor = Color.FromArgb(0, 225, 255);
                    break;
                case "G":
                    //Counting = false;
                    State_Ready = false;
                    label_Time_display.BackColor = Color.FromArgb(128, 255, 128);
                    Write_dataGridView(DataGvRowInd, DataGvColInd, label_Time_display.Text);
                    SaveToExcel(ExcelFilePath, xlCells_RowInd, xlCells_ColInd, label_Time_display.Text);
                    DataGvColInd += 1; xlCells_ColInd += 1;
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
                    label_Time_display.Text = "Fail";
                    Write_dataGridView(DataGvRowInd, DataGvColInd, "XX:XX.XXX");
                    SaveToExcel(ExcelFilePath, xlCells_RowInd, xlCells_ColInd, label_Time_display.Text);
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
                        stopWatch.Start();
                        Int32 length = My_SerialPort.Read(buffer, 0, buffer.Length);
                        Array.Resize(ref buffer, length);

                        string buf = Encoding.UTF8.GetString(buffer);
                        string[] StrArr = buf.Split(new string[] { "\r\n" }, StringSplitOptions.None);

                        bool EnDisplay = true;
                        foreach (string str in StrArr)
                        {

                            // for 60Hz display
                            if (str.Length == 8 && EnDisplay)
                            {

                                EnDisplay = false;
                                Display d = new Display(ShowTime);
                                string time = Format_MilliSecond(HexToUint(str));
                                this.Invoke(d, new Object[] { time });
                            }
                            else if (str.Length == 1)
                            {
                                Control c = new Control(Command);
                                this.Invoke(c, new Object[] { str });
                            }

                        }
                        Array.Resize(ref buffer, 1024);
                        stopWatch.Stop();
                        Console.WriteLine(stopWatch.ElapsedMilliseconds);
                    }
                    Thread.Sleep(40);
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
        private void PorSelector_DropDown(object sender, EventArgs e)
        {
            PorSelector.Items.Clear();
            var searcher = new ManagementObjectSearcher("SELECT DeviceID,Caption FROM WIN32_SerialPort");
            foreach (ManagementObject port in searcher.Get())
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
            ComPortName.SelectedIndex = PorSelector.SelectedIndex;
            Console_Connect(ComPortName.Text, 115200);
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
        private void button_Sand_Click(object sender, EventArgs e)
        {
            SendString(textBoxSend.Text);
            textBoxSend.Clear();
        }
        private void button_Clear_Click(object sender, EventArgs e)
        {
            textBoxReceive.Clear();
        }
        private void button_Command_Ready_Click(object sender, EventArgs e)
        {
            if (State_Ready == false)
            {
                SendString("R\n");
            }
            //My_SerialPort.Write("R\n");
        }
        private void button_Command_Fail_Click(object sender, EventArgs e)
        {
            if (State_Failing == false)
            {
                SendString("F\n");
                //My_SerialPort.Write("F\n");
            }
        }
        private void button_Command_Restart_Click(object sender, EventArgs e)
        {
            Reset();
            //My_SerialPort.Write("S\n");
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
        private void FrontPage_Load(object sender, EventArgs e)
        {
            //string strPath = System.Environment.CurrentDirectory;
            //string strPath = System.Windows.Forms.Application.StartupPath;//啟動路徑
            string strPath = System.Windows.Forms.Application.ExecutablePath;
            //label_excel_1.Text = strPath;

            //畫面開啟時直接連接Com10
            //Console_Connect("COM10", 115200);
        }
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
        public void Write_dataGridView(int GwRows, int GwColumns, string GwScore)
        {
            if (DataGvLoaded)
                dataGridView1.Rows[GwRows].Cells[GwColumns].Value = GwScore;
        }
        private void button_excel_1_Click(object sender, EventArgs e)
        {
            Create_Excel_template();
        }
        private void button_excel_2_Click(object sender, EventArgs e)
        {
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
        }
        private void button_excel_3_Click(object sender, EventArgs e)
        {
            SaveToExcel(ExcelFilePath, xlCells_RowInd, xlCells_ColInd, textBoxSend.Text);
            Write_dataGridView(DataGvRowInd, DataGvColInd, textBoxSend.Text);
        }
        private void button_excel_4_Click(object sender, EventArgs e)
        {
            Excel.Application oXL;
            Excel._Workbook oWB;
            string path = ExcelFilePath;
            //Start Excel and get Application object.
            oXL = new Excel.Application();
            //Get a new workbook.
            oWB = oXL.Workbooks.Open(path);
            oXL.Visible = true;
            oXL.UserControl = true;
        }
        private void button_excel_5_Click(object sender, EventArgs e)
        {
            getExcelFile(ExcelFilePath);
        }

        private void button_Command_3_Click(object sender, EventArgs e)
        {
            
        }

        private void button_Command_4_Click(object sender, EventArgs e)
        {
            
        }

        private void button_Inf_Previous_Click(object sender, EventArgs e)
        {
            if (DataGvLoaded == true)
            {
                DataGvRowInd += -1;
                DataGv_Get_Rows_Location = false;
                dataGridView1.Focus();
                dataGridView1.CurrentCell = dataGridView1[DataGvColInd, DataGvRowInd];
                dataGridView1.BeginEdit(true);
                DataGv_Get_Rows_Location = true;
            }
        }

        private void button_Inf_Next_Click(object sender, EventArgs e)
        {
            if (DataGvLoaded == true)
            {
                DataGvRowInd += 1;
                DataGv_Get_Rows_Location = false;
                dataGridView1.Focus();
                dataGridView1.CurrentCell = dataGridView1[DataGvColInd, DataGvRowInd];
                dataGridView1.BeginEdit(true);
                DataGv_Get_Rows_Location = true;
            }
        }

        private void button_Round_Previous_Click(object sender, EventArgs e)
        {
            if (DataGvLoaded == true)
            {
                DataGvColInd += -1;
                DataGv_Get_Current_Location = false;
                dataGridView1.Focus();
                dataGridView1.CurrentCell = dataGridView1[DataGvColInd, DataGvRowInd];
                dataGridView1.BeginEdit(true);
                DataGv_Get_Current_Location = true;
            }
        }

        private void button_Round_Next_Click(object sender, EventArgs e)
        {
            if (DataGvLoaded == true)
            {
                DataGvColInd += 1;
                DataGv_Get_Current_Location = false;
                dataGridView1.Focus();
                dataGridView1.CurrentCell = dataGridView1[DataGvColInd, DataGvRowInd];
                dataGridView1.BeginEdit(true);
                DataGv_Get_Current_Location = true;
            }
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

                if (DataGvColInd > (4 + TotalRuns - 1)) DataGvColInd = (4 + TotalRuns - 1);
                else if (DataGvColInd < 4) DataGvColInd = 4;
                if (DataGvRowInd > totalrowCount) DataGvRowInd = totalrowCount;
                else if (DataGvRowInd < 0) DataGvRowInd = 0;
                xlCells_RowInd = DataGvRowInd + 2;
                xlCells_ColInd = DataGvColInd + 1;
                textBox_Round.Text = "Rund\r\n "+ (DataGvColInd - 3);
                if (dataGridView1.Rows[DataGvRowInd].Cells[2].Value != null && dataGridView1.Rows[DataGvRowInd].Cells[3].Value != null)
                {
                    textBox_Team_Information.Text = dataGridView1.Rows[DataGvRowInd].Cells[2].Value.ToString() +
                        "\r\n" + dataGridView1.Rows[DataGvRowInd].Cells[3].Value.ToString();
                    //F2.Team_Information = dataGridView1.Rows[DataGvRowInd].Cells[2].Value.ToString() +
                    //    "\r\n" + dataGridView1.Rows[DataGvRowInd].Cells[3].Value.ToString();
                    //F2.Round = textBox_Round.Text;
                    //label_excel_2.Text = 
                    ////    dataGridView1.Rows[DataGvRowInd].Cells[2].Value.ToString() + "           " + 
                    ////    dataGridView1.Rows[DataGvRowInd].Cells[3].Value.ToString() + "                      " +
                    //    dataGridView1.Columns[DataGvColInd].HeaderText + "           " +
                    //    "Row" + DataGvRowInd.ToString() +
                    //    "Column" + DataGvColInd.ToString();
                }
            }
        }
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
        
    }
}
