using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using TimerLibrary;
using DataTable = System.Data.DataTable;

namespace WindowsFormsApp1
{
    public partial class Form3 : Form
    {
        //public Competition Team_List = new Competition();

        public Form3()
        {
            InitializeComponent();

        }
        public string FilePath;
        private void button1_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Excel 活頁簿 (*.xlsx)|*.xlsx|Excel 97-2003 (*.xls)|*.xls|文字檔 (Tab 字元分隔) (*.txt)|*.txt";
                ofd.Title = "Select Excel file";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    FilePath = ofd.FileName;
                    DataTable dt = null;
                    var pakge = new ExcelPackage(FilePath);
                    ExcelWorkbook workbook = pakge.Workbook;
                    if (workbook != null)
                    {
                        ExcelWorksheet worksheet = workbook.Worksheets.First();
                        dt = WorksheetToTable(worksheet);
                        dataGridView1.DataSource = dt;
                    }
                }
            }
        }
        protected override bool ProcessCmdKey(ref System.Windows.Forms.Message msg, System.Windows.Forms.Keys keyData) //激活回车键
        {
            int WM_KEYDOWN = 256;
            int WM_SYSKEYDOWN = 260;

            if (msg.Msg == WM_KEYDOWN | msg.Msg == WM_SYSKEYDOWN)
            {
                switch (keyData)
                {
                    case Keys.Escape:
                        this.Close();//csc关闭窗体
                        break;

                    case Keys.F12:
                        MessageBox.Show("F12");
                        break;

                }

            }
            return false;
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
                for (int j =1; j <= cols; j++)
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

        private void Form3_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyCode)
            {
                case Keys.F11:
                    //if (FormBorderStyle == FormBorderStyle.None)
                    //{
                    //    this.FormBorderStyle = FormBorderStyle.Sizable;  //設定窗體為無邊框樣式
                    //    this.WindowState = FormWindowState.Normal;       //最大化窗體
                    //}
                    //else
                    //{
                    //    this.FormBorderStyle = FormBorderStyle.None;     //設定窗體為無邊框樣式
                    //    this.WindowState = FormWindowState.Maximized;    //最大化窗體
                    //}
                    MessageBox.Show("F11");
                    break;
                case Keys.Enter:
                    MessageBox.Show("Enter");
                    break;
                default:
                    break;
            }
        }

        private void Form3_Load(object sender, EventArgs e)
        {
            Competition Team_List = new Competition();
            Team_List.Team.Add(new TEAM("t1"));
            Team_List.Team[0].Time.Add(new TimerData(154612));
            Team_List.Team[0].Time.Add(new TimerData(74554));
            Team_List.Team[0].Time.Add(new TimerData(7867864));
            Team_List.Team[0].ID = "A1U-054";
            Team_List.Team[0].Oeder = "1";
            Team_List.Team[0].Name = "T1";
            Console.WriteLine(Team_List.Team[0].ID);
            Console.WriteLine(Team_List.Team[0].Oeder);
            Console.WriteLine(Team_List.Team[0].Name);
            Console.WriteLine("\nBefore sort:");
            foreach (var aPart in Team_List.Team[0].Time)
            {
                Console.WriteLine(aPart);
            }
            Console.WriteLine(Team_List.Team[0].Time[0]);
            Team_List.Team[0].Time.Sort();
            Console.WriteLine("\nAfter sort by part number:");
            foreach (var aPart in Team_List.Team[0].Time)
            {
                Console.WriteLine(aPart);
            }
            /*
            TEAM data = new TEAM("87");
            data.Time.Add(new TimerData(1234));
            data.Time.Add(new TimerData(1000));
            data.Time.Add(new TimerData(1555));
            data.ID = "A1U-005";
            Console.WriteLine(data.ID);
            Console.WriteLine("\nBefore sort:");
            foreach (var aPart in data.Time)
            {
                Console.WriteLine(aPart);
            }
            Console.WriteLine();
            // Call Sort on the list. This will use the
            // default comparer, which is the Compare method
            // implemented on Part.
            data.Time.Sort();
            Console.WriteLine(data.Time.Max());
            Console.WriteLine(data.Time.Min());
            Console.WriteLine();

            Console.WriteLine("\nAfter sort by part number:");
            foreach (var aPart in data.Time)
            {
                Console.WriteLine(aPart);
            }
            */
        }
    }
}
