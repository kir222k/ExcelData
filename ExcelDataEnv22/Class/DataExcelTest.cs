using System;
using System.Collections.Generic;
using System.Linq;
//using System.Data;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.IO;
using System.Windows.Forms;
using ExcelData.Model;
using ExcelData.Class;
//#error version

namespace ExcelData.Class
{
    public static class DataExcelTest
    {
        public static string[,] GetDataExcelToArrayTest(string fileExcelName, string sheetExcelName, int totalRows=-1 , int totalColumns=-1)
        {
            ExcelPackage excelFile = new ExcelPackage(new FileInfo(fileExcelName));
            ExcelWorksheet worksheet = excelFile.Workbook.Worksheets[sheetExcelName];

            var DE = new DataExcel();


            if ((totalRows == -1) || (totalColumns == -1))
            {
                totalRows = worksheet.Dimension.End.Row+1;
                totalColumns = worksheet.Dimension.End.Column+1;
            }


            string[,] excelTable = new string[totalRows, totalColumns];

            for (int i = 0; i < totalRows -1; i++)
            {
                for (int j = 0; j < totalColumns -1; j++)
                {
                    ExcelRangeBase Cell = worksheet.Cells[i + 1, j + 1];
                    if (Cell.Value != null)
                    {

                        //// проверяем , если это число - округляем до 2 знаков 
                        string strChek = Convert.ToString(Cell.Value);

                        if (
                            (Cell.Style.Numberformat.Format.Contains("0.0")) ||
                            (ValueChek.IsDigitStr(strChek))
                           )
                        {
                            double valueDouble = Convert.ToDouble(ValueChek.GetPointValid(Convert.ToString(Cell.Value)));

                            string str = Convert.ToString(Math.Round(valueDouble, Const.RoundForDouble));
                            str = ValueChek.GetAddZeroStr(str, Const.RoundForDouble);

                            excelTable[i, j] = "Value=" + str + " Format=" + Cell.Style.Numberformat.Format;
                        }
                        else
                            excelTable[i, j] = "Value=" + Convert.ToString(Cell.Value) + " Format=" + Cell.Style.Numberformat.Format;

                        //excelTable[i, j] = Convert.ToString(Cell.Value); // + " Type = " + Cell.Style.Numberformat.Format);

                    }
                    else
                    {
                        excelTable[i, j] = Const.NullTextReplace;
                    }
                }
            }
            return excelTable;
        }


        public static void TestDial()
        {
            OpenFileDialog dialog = new OpenFileDialog();
            DialogResult res = dialog.ShowDialog();
        }
    }
}
