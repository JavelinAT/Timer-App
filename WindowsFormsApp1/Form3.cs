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
    }
}
