using OfficeOpenXml;
using SpreadsheetLight;
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

namespace ImportExcelToDataGrid
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FrmImportExcelToDatagrid frmImportExcelToDataGrid = new FrmImportExcelToDatagrid();
            frmImportExcelToDataGrid.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            String strTest = Directory.GetCurrentDirectory();
            String path = Path.Combine(Directory.GetCurrentDirectory(), @"FileExcel/ImportExcelToDatagrid.xlsx");
            //this.readXLS(path);

            using (var package = new ExcelPackage(path))
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                var sheet = package.Workbook.Worksheets[0];
                var persons = this.GetList<PersonasViewModel>(sheet);
            }
        }

        public void readXLS(string FilePath)
        {
            FileInfo existingFile = new FileInfo(FilePath);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                //get the first worksheet in the workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int colCount = worksheet.Dimension.End.Column;  //get Column Count
                int rowCount = worksheet.Dimension.End.Row;     //get row count
                for (int row = 1; row <= rowCount; row++)
                {
                    for (int col = 1; col <= colCount; col++)
                    {
                        Console.WriteLine(" Row:" + row + " column:" + col + " Value:" + worksheet.Cells[row, col].Value?.ToString().Trim());
                    }
                }
            }
        }

        private List<T> GetList<T>(ExcelWorksheet sheet)
        {
            List<T> darrList = new List<T>();
            var columnInfo = Enumerable.Range(1, sheet.Dimension.Columns).ToList().Select(n =>
                new { Index = n, ColumnName = sheet.Cells[1, n].Value.ToString() }
                );

            for (int intRow = 2; intRow < sheet.Dimension.Rows; intRow++)
            {
                T obj = (T)Activator.CreateInstance(typeof(T));
                foreach (var pro in typeof(T).GetProperties())
                {
                    int intCol = columnInfo.SingleOrDefault(c => c.ColumnName == pro.Name).Index;
                    var val = sheet.Cells[intRow, intCol].Value;
                    var propType = pro.PropertyType;
                    pro.SetValue(obj, Convert.ChangeType(val, propType));
                }
                darrList.Add(obj);
            }
            return darrList;
        }

        
    }
}
