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
using OfficeOpenXml;
namespace WindowsFormsApp1
{
    public partial class Form3 : Form
    {
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
    }
}
