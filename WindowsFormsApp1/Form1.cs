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


using System.Diagnostics;

namespace WindowsFormsApp1
{
    public partial class FrontPage : Form
    {
        private SerialPort My_SerialPort;
        private bool Console_receiving = false;
        private bool Counting = false;
        private Thread t;
        private string ExcelFilePath;
        Flag flag = new Flag(false, false, false);

        public struct Flag
        {
            public bool Ready;
            public bool Failing;
            public bool Counting;
            public Flag(bool r, bool f, bool c)
            {
                this.Ready = r;
                this.Failing = f;
                this.Counting = c;
            }
        }
        delegate void Display(string buffer);//1    //使用委派顯示
        public FrontPage()
        {
            InitializeComponent();
        }
        public void ConsoleShow(string buffer)//3
        {
            //textBoxReceive.Text = buffer + "\r\n";
            string[] StrArr = buffer.Split(new string[] { "\r\n" }, StringSplitOptions.None);
            foreach (string str in StrArr)
            {
                if (str.Length == 8)
                {
                        label_display.Text = Format_MilliSecond(HexToUint(str));
                }
                else if(str.Length == 1)
                {
                    textBoxReceive.Text += str + "\r\n";
                    textBoxReceive.SelectionStart = textBoxReceive.Text.Length;
                    textBoxReceive.ScrollToCaret();
                    switch (str)
                    {
                        case "S":
                            Counting = true;
                            label_display.BackColor = Color.FromArgb(0, 225, 255);
                            break;
                        case "G":
                            Counting = false;
                            flag.Ready = false;
                            label_display.BackColor = Color.FromArgb(128, 255, 128);
                            label1.Text +=label_display.Text + "\r\n";
                            break;
                        case "C":
                            label_display.BackColor = Color.Transparent;
                            break;
                        case "R":   //Ready
                            flag.Ready = true;
                            flag.Failing = false;
                            label_display.BackColor = Color.FromArgb(128, 255, 128);
                            label_display.Text = "Ready";
                            break;
                        case "F":   //Fail
                            flag.Failing = true;
                            flag.Ready = false;
                            label_display.BackColor = Color.FromArgb(192, 0, 0);
                            label_display.Text = "Fail";
                            label1.Text += "XX:XX.XXX\r\n";
                            break;
                        
                    }
                }
            }
            /*
            System.Console.WriteLine(buffer);
            textBoxReceive.Text += buffer + "\r\n";
            /*
            //滾動至最下方
            textBoxReceive.SelectionStart = textBoxReceive.Text.Length;
            textBoxReceive.ScrollToCaret();
            */
        }
        public void Reset()
        {
            flag.Ready = false;
            flag.Failing = false;
            label_display.Text = "00:00.000";
            label_display.BackColor = Color.Transparent;
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
                    t = new Thread(DoReceive);
                    t.IsBackground = true;
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
            uint data = 0;

            for (int i = 0; i < hexString.Length; i = i + 2)
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
                        Int32 length = My_SerialPort.Read(buffer, 0, buffer.Length);
                        Array.Resize(ref buffer, length);

                        string buf = Encoding.UTF8.GetString(buffer);

                        Display d = new Display(ConsoleShow);//2
                        this.Invoke(d, new Object[] { buf });

                        Array.Resize(ref buffer, 1024);
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
        private void PorSelector_DropDown(object sender, EventArgs e)
        {
            PorSelector.Items.Clear();
            PorSelector.Items.AddRange(SerialPort.GetPortNames());
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
            Console_Connect(PorSelector.Text, 115200);
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
            if (flag.Ready == false)
            {
                SendString("R\n");
            }
            //My_SerialPort.Write("R\n");
        }
        private void button_Command_Fail_Click(object sender, EventArgs e)
        {
            if(flag.Failing == false)
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
            //string str = System.Environment.CurrentDirectory;
            //string str = System.Windows.Forms.Application.StartupPath;//啟動路徑
            string strPath = System.Windows.Forms.Application.ExecutablePath;
            label_excel_1.Text = strPath;
            //Form2 frm = new Form2();
            //frm.Show(this);

            //畫面開啟時直接連接Com10
            //Console_Connect("COM10", 115200);
        }
        private void OpenExcel()
        {
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;
            Excel.Range oRng;

            try
            {
                //Start Excel and get Application object.
                oXL = new Excel.Application();
                oXL.Visible = true;

                //Get a new workbook.
                oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
                oSheet = (Excel._Worksheet)oWB.ActiveSheet;

                //Add table headers going cell by cell.
                oSheet.Cells[1, 1] = "First Name";
                oSheet.Cells[1, 2] = "Last Name";
                oSheet.Cells[1, 3] = "Full Name";
                oSheet.Cells[1, 4] = "Salary";

                //Format A1:D1 as bold, vertical alignment = center.
                oSheet.get_Range("A1", "D1").Font.Bold = true;
                oSheet.get_Range("A1", "D1").VerticalAlignment =
                Excel.XlVAlign.xlVAlignCenter;

                // Create an array to multiple values at once.
                string[,] saNames = new string[5, 2];

                saNames[0, 0] = "John";
                saNames[0, 1] = "Smith";
                saNames[1, 0] = "Tom";
                saNames[1, 1] = "Brown";
                saNames[2, 0] = "Sue";
                saNames[2, 1] = "Thomas";
                saNames[3, 0] = "Jane";
                saNames[3, 1] = "Jones";
                saNames[4, 0] = "Adam";
                saNames[4, 1] = "Johnson";

                //Fill A2:B6 with an array of values (First and Last Names).
                oSheet.get_Range("A2", "B6").Value2 = saNames;

                //Fill C2:C6 with a relative formula (=A2 & " " & B2).
                oRng = oSheet.get_Range("C2", "C6");
                oRng.Formula = "=A2 & \" \" & B2";

                //Fill D2:D6 with a formula(=RAND()*100000) and apply format.
                oRng = oSheet.get_Range("D2", "D6");
                oRng.Formula = "=RAND()*100000";
                oRng.NumberFormat = "$0.00";

                //AutoFit columns A:D.
                oRng = oSheet.get_Range("A1", "D1");
                oRng.EntireColumn.AutoFit();

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
        private void SaveOnExcel(string path)
        {
            //string FileStr = "D:\\test";
            Excel.Application oXL;
            Excel._Workbook oWB;
            Excel._Worksheet oSheet;
            //Excel.Range oRng;
            

            try
            {
                //Start Excel and get Application object.
                oXL = new Excel.Application();
                oXL.Visible = false;

                //Get a new workbook.
                //oWB = (Excel._Workbook)(oXL.Workbooks.Add(Missing.Value));
                oWB = oXL.Workbooks.Open(path);
                oSheet = oWB.Worksheets[1];

                //Add table headers going cell by cell.
                oSheet.Cells[1, 1] = "Team";
                oSheet.Cells[1, 2] = "Name";
                oSheet.Cells[1, 3] = "Score";

                //Format A1:D1 as bold, vertical alignment = center.
                oSheet.get_Range("A1", "C1").Font.Bold = true;
                oSheet.get_Range("A1", "C1").VerticalAlignment =
                Excel.XlVAlign.xlVAlignCenter;

                string[] StrArr = label1.Text.Split(new string[] { "\r\n" }, StringSplitOptions.None);
                int count = 1;
                foreach (string str in StrArr)
                {
                    count += 1;
                    oSheet.Cells[count, 3].NumberFormat = "@";
                    oSheet.Cells[count, 3].Value = str;
                }
                
                oXL.Visible = false;
                //oXL.UserControl = true;
                oWB.Save();
                oWB.Close();
                oXL.Quit();
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
        public void getExcelFile()
        {

            //Create COM Objects. Create a COM object for everything that is referenced
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(label_excel_1.Text);
            Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlWorksheet.UsedRange;

            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;
            //--------------------------------------

            textBox1.Text = null;
            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {
                    if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                    {
                        string strdata = xlRange.Cells[i, j].Value2.ToString();
                        if (i == 1)
                        {
                            dataGridView1.Columns.Add("Col1", strdata);
                        }
                        else
                            this.dataGridView1.Rows[i - 2].Cells[j - 1].Value = strdata;
                    }
                }
                if(i != rowCount) this.dataGridView1.Rows.Add();
            }
            //--------------------------------------
            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!

            //--------------------------------------
            //------源code

            //textBox1.Text = null;
            //for (int i = 1; i <= rowCount; i++)
            //{
            //    for (int j = 1; j <= colCount; j++)
            //    {
            //        //new line
            //        if (j == 1)
            //        {
            //            Console.Write("\r\n");
            //            textBox1.Text += "\r\n";
            //        }

            //        //write the value to the console
            //        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
            //        {
            //            Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");
            //            textBox1.Text += xlRange.Cells[i, j].Value2.ToString() + "\t";
            //        }
            //        else textBox1.Text += "\t";
            //    }
            //}
            //--------------------------------------

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
        }
        private void button_excel_1_Click(object sender, EventArgs e)
        {
            OpenExcel();
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
            SaveOnExcel(label_excel_1.Text);
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
            getExcelFile();
            //DataTable TableValue = ImportExcel("Table1");
            //dataGridView1.DataSource = TableValue;
        }
        public DataTable ImportExcel(string SheetName)
        {
            
            DataTable dataTable = new DataTable();

            

                //2.提供者名稱  Microsoft.Jet.OLEDB.4.0適用於2003以前版本，Microsoft.ACE.OLEDB.12.0 適用於2007以後的版本處理 xlsx 檔案
                string ProviderName = "Microsoft.ACE.OLEDB.12.0;";

                //3.Excel版本，Excel 8.0 針對Excel2000及以上版本，Excel5.0 針對Excel97。
                string ExtendedString = "'Excel 8.0;";

                //4.第一行是否為標題(;結尾區隔)
                string HDR = "No;";

                //5.IMEX=1 通知驅動程序始終將「互混」數據列作為文本讀取(;結尾區隔,'文字結尾)
                string IMEX = "0';";

                //=============================================================
                //連線字串
                string connectString =
                        "Data Source=" + ExcelFilePath + ";" +
                        "Provider=" + ProviderName +
                        "Extended Properties=" + ExtendedString +
                        "HDR=" + HDR +
                        "IMEX=" + IMEX;
                //=============================================================

                using (OleDbConnection Connect = new OleDbConnection(connectString))
                {
                    Connect.Open();
                    string queryString = "SELECT * FROM [" + SheetName + "$]";
                    try
                    {
                        using (OleDbDataAdapter dr = new OleDbDataAdapter(queryString, Connect))
                        {
                            dr.Fill(dataTable);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("異常訊息:" + ex.Message, "異常訊息");
                    }
                }
            return dataTable;
        }

        private void button_Command_3_Click(object sender, EventArgs e)
        {

        }

        private void button_Command_4_Click(object sender, EventArgs e)
        {

        }
    }
}
