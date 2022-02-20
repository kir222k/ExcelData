/* Кирилл Уваров 2022г. 10 февраля. u.k.send@gmail.com. +79062644029
*/

using System;
using ExcelData;
using System.Diagnostics;
using ExcelData.Class;
using ExcelData.Sys;
using ExcelData.Model;
using System.Collections.Generic;
using System.Linq;

namespace TestData
{
    class Program
    {
        static void Main()
        {
            var sw = new Stopwatch();
            sw.Start();

            // Metod2EpPlus();
            // Metod3EpPlus();

             Metod4EpPlus();

            Console.WriteLine(ValueChek.IsDigitStr("0.3333333"));

            sw.Stop();

            Console.WriteLine($"\nВремя выполнения:\n=>" +
                $"\n{sw.Elapsed} секунд   " +
                $"{sw.ElapsedMilliseconds} миллисекунд");
            Console.WriteLine("\nВыход..");
        }


        static void Metod3EpPlus()
        {
            DataExcel ED = new DataExcel(Const.FileXlsName, Const.ExcelWorksheet);
            PullPushData PP = new PullPushData(ED.GetDataExel().Array);
            Console.WriteLine(PP);
        }

        /// <summary>
        /// Вывод в консоль данных листа, тест на тип значения
        /// </summary>
        static void Metod4EpPlus ()
        {
            //         public static string[,] GetDataExcelToArrayTest(string fileExcelName, string sheetExcelName, int totalRows, int totalColumns)
            int totalRows = -1;
            int totalColumns = -1;

            string[,] arrayData;//= new string[totalRows, totalColumns];

            Console.WriteLine(Const.FileXlsName);

            arrayData = DataExcelTest.GetDataExcelToArrayTest(Const.FileXlsName, Const.ExcelWorksheet, totalRows, totalColumns);

            int rows = arrayData.GetUpperBound(0) + 1;    // количество строк
            int columns = arrayData.GetUpperBound(1) + 1; // количество столбцов

            for (int i = 0; i < rows - 1; i++)
            {
                for (int j = 0; j< columns - 1; j++)
                {
                    Console.Write("i=" + i+",j=" + j.ToString() + " <" + arrayData[i, j] + "> ");
                }
                Console.WriteLine("");
            }
        }


        //работает мега быстро
        static void Metod2EpPlus()
    {
        DataExcel ED = new DataExcel(Const.FileXlsName, Const.ExcelWorksheet);
        //DataExcel ED = new DataExcel("", Const.ExcelWorksheet);
        // DataExcel ED = new DataExcel(Const.FileXlsName, "");
        var Result  = new ArrayWithComments();
        Result = ED.GetDataExel();

        if (Result.Array != null)
        {
            // Сам массив данных
            string[,] strTable = Result.Array;
            // количество строк
            int rows = strTable.GetUpperBound(0) + 1;
            // количество столбцов
            int columns = strTable.GetUpperBound(1) + 1;      
            // Создадим класс для получ. данных 
            var PP = new PullPushData(strTable);

            // ТЕСТ СПИСКА БЛОКОВ Вывод данных в консоль
            // список из Блок.i.j
            var sPP = PP.GetExcelRangeBlock();
            string str = "";
            LogEasy.DeleteFileLog(Const.LogFileName);
            foreach (var blData in sPP)
            {
                str += "\n\nБлок: " + blData.TextValue;
                str += "\nКоординаты_ячейки=> ";
                str += "\nСтрока= " + blData.RowCell + "\nСтолбец= " + blData.ColumnCell;

                Console.WriteLine(str);

                LogEasy.WriteLog(str, Const.LogFileName);
                str = "";
            }
            Console.WriteLine($"Строк={rows}  Столбцов={columns}");
            Console.WriteLine("\n\nКоличество элем-тов в списке блоков= " + sPP.Count.ToString() + "\n");


            // Вывод табл данных в файл
            LogEasy.DeleteFileLog(Const.LogFileTable);
            for (int i = 0; i < rows - 1; i++)
            {
                string st = "";
                for (int j = 0; j < columns - 1; j++)
                {
                        st += strTable[i, j] + ";";
                }
                LogEasy.WriteLog(st, Const.LogFileTable);
            }

        }

        Console.WriteLine("\nResult.Comments:\n=>\n" + Result.Comments);
    }

    }
}
