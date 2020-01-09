using System;
using System.IO;
using System.Reflection;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace Read_in_Excel.Models
{
    public class ReadExcelModel
    {

        public void Read()
        {
            Microsoft.Office.Interop.Excel.Application _excelApp = new Microsoft.Office.Interop.Excel.Application();
            //_excelApp.Visible = true;
            String filename = @"C:\Users\$ARVARBEK\Desktop\Shablon.xlsx";
            Workbook workbook = _excelApp.Workbooks.Open(filename);
            Worksheet worksheet = (Worksheet)workbook.Worksheets[1];
            Range excelRange = worksheet.UsedRange;
            object[,] valueArray = (object[,])excelRange.get_Value(XlRangeValueDataType.xlRangeValueDefault);
            try
            {
                for (int row = 1; row <= worksheet.UsedRange.Rows.Count; ++row)
                {
                    for (int col = 1; col <= worksheet.UsedRange.Columns.Count; ++col)
                    {
                        valueArray[row, col].ToString();
                    }
                }
            }
            catch (Exception) { }
            workbook.Close(false);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
            _excelApp.Quit();
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(_excelApp);            
        }        
    }
}