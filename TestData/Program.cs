﻿using System;
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
            try
            {

            // var pr = ExcelBook.ExcelBookTest();
            // Console.WriteLine(pr);

            //Excel.Worksheet ws;
            //ws = ExcelBook.GetExcelSheet(Const.FileXlsName, Const.ExcelWorksheet);
            //Console.WriteLine("\nКнига получена");
            Console.WriteLine("\nПолучение данных");
            ExcelBook ExBook = new ExcelBook();

            int x = 150; int y = 80; string range = "A1";
            var Arr = ExBook.GetArrayBasedCell(Const.FileXlsName, Const.ExcelWorksheet , x, y, range);
            Console.WriteLine("\nМассив получен");

            for (int i = 0; i < x; i++)
            {
                for (int j = 0; j < y; j++)
                {
                    Console.Write(Arr[i, j].ToString() + " ");
                }
                Console.WriteLine("\n");
            }
            // ws.Application.Quit();

            /*

            // 3. Получим массив 50*50 значений  от А1, поищем там ячейку с текстом "Блок"
            int x1 = x; int y1 = y; string range1 = range;
            //ws = ExcelBook.GetExcelSheet(Const.FileXlsName, Const.ExcelWorksheet);
            //var ArrForBlock = ExcelBook.GetArrayBasedCell(ws, x1, y1, range1);
            var ArrForBlock = Arr;
            for (int i = 0; i < x1; i++)
            {
                for (int j = 0; j < y1; j++)
                {
                    // Console.Write(Arr[i, j].ToString() + " ");
                    if (ArrForBlock[i, j] == "Блок")
                    {
                        // кортеж с данными (строка, столбец) , где сидит слово "Блок"
                        var rangeBlock = (i, j); // столбец, в кот. будут имена блоков найден
                        Console.WriteLine("\nСЛОВО: Блок =>" + "Строка: "  + rangeBlock.i + " Столбец: " + rangeBlock.j);
                        //Console.WriteLine("\nСЛОВО: Блок =>" +  ws.Range();
                        break;
                    }
                }
            }
            */

            // ws.Application.Quit();

            Console.WriteLine("\n\nДля завершения работы нажмите л. кл.");
            Console.ReadLine();
            }
            catch (Exception)
            {

                throw;
            }
        }
    }
}
