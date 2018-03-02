using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace DevCenterBot
{
    public class ExcelHelper
    {
        public void ExportExcel(List<ExcelUserVM> input)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = new Excel.Application();

            //Create an Excel workbook instance and open it from the predefined location
            Excel.Workbook excelWorkBook = (Excel.Workbook)(excelApp.Workbooks.Add(Missing.Value));
            Excel.Worksheet excelWorkSheet = excelWorkBook.Sheets.Add();
            excelWorkSheet.Name = "Extracted Details";
            for (int rows = 1; rows <= input.Count; rows++)
            {
                if (rows == 1)
                {
                    excelWorkSheet.Cells[1, 1] = "Name";
                    excelWorkSheet.Cells[1, 2] = "Following Count";
                }
                
                excelWorkSheet.Cells[(rows + 1), 1] = input[rows].Name;
                excelWorkSheet.Cells[(rows + 1), 2] = input[rows].FollowerCount;
                
            }
            excelWorkBook.SaveAs(System.IO.Directory.GetCurrentDirectory() + @"\DcUsers.xlsx", Excel.XlFileFormat.xlOpenXMLWorkbook, Missing.Value,
                                    Missing.Value, false, false, Excel.XlSaveAsAccessMode.xlNoChange,
                                    Excel.XlSaveConflictResolution.xlUserResolution, true,
                                    Missing.Value, Missing.Value, Missing.Value);
            excelWorkBook.Close();
            excelApp.Quit();
        }        
    }

    public class ExcelUserVM
    {
        public string Name { get; set; }
        public int FollowerCount { get; set; }
    }
}
