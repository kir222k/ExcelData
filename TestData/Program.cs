using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelData;
using Excel = Microsoft.Office.Interop.Excel;

namespace TestData
{
    class Program
    {
        static void Main(string[] args)
        {
            var pr = ExcelBook.ExcelBookTest();
            Console.WriteLine(pr);

            int x = 10; int y = 2; string range = "C16";
            var Arr = ExcelBook.GetArrayBasedCell(
                ExcelBook.GetExcelSheet(Const.FileXlsName, Const.ExcelWorksheet), x, y, range);

            for (int i = 0; i < x; i++)
            {
                for (int j = 0; j < y; j++)
                {
                    Console.Write(Arr[i, j].ToString() + " ");
                }
                Console.WriteLine("\n");
            }
        

            Console.ReadKey();
        }
    }
}
