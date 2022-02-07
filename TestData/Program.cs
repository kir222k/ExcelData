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
            Metod3EpPlus();

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

            /*
            var listAtts = new List<ExcelRangeText>();
            listAtts = PP.GetExcelRangeAttribute();
            string strAtts = "\nАтрибуты в заголовке:\n=>";
            foreach (ExcelRangeText att in listAtts)
            {
                strAtts += "\n" + att.TextValue;
            }
            Console.WriteLine(strAtts);
            */
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

            // ТЕСТ СПИСКА БЛОК-АТРИБУТЫ Вывод данных в консоль
            // 






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

        // работает очень медленно
        #region Metod1
        /*
        static void Metod1 ()
        {
            var pr = ExcelBook.ExcelBookTest();
            Console.WriteLine(pr);

                // var pr = InteropLinktoExcel.ExcelBookTest();
                // Console.WriteLine(pr);

                //Excel.Worksheet ws;
                //ws = InteropLinktoExcel.GetExcelSheet(Const.FileXlsName, Const.ExcelWorksheet);
                //Console.WriteLine("\nКнига получена");
                Console.WriteLine("\nПолучение данных");
                InteropLinktoExcel ExBook = new InteropLinktoExcel();

                int x = 150; int y = 80; string range = "A1";
                var Arr = ExBook.GetArrayBasedCell(Const.FileXlsName, Const.ExcelWorksheet, x, y, range);
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

                

                //// 3. Получим массив 50*50 значений  от А1, поищем там ячейку с текстом "Блок"
                //int x1 = x; int y1 = y; string range1 = range;
                ////ws = InteropLinktoExcel.GetExcelSheet(Const.FileXlsName, Const.ExcelWorksheet);
                ////var ArrForBlock = InteropLinktoExcel.GetArrayBasedCell(ws, x1, y1, range1);
                //var ArrForBlock = Arr;
                //for (int i = 0; i < x1; i++)
                //{
                //    for (int j = 0; j < y1; j++)
                //    {
                //        // Console.Write(Arr[i, j].ToString() + " ");
                //        if (ArrForBlock[i, j] == "Блок")
                //        {
                //            // кортеж с данными (строка, столбец) , где сидит слово "Блок"
                //            var rangeBlock = (i, j); // столбец, в кот. будут имена блоков найден
                //            Console.WriteLine("\nСЛОВО: Блок =>" + "Строка: "  + rangeBlock.i + " Столбец: " + rangeBlock.j);
                //            //Console.WriteLine("\nСЛОВО: Блок =>" +  ws.Range();
                //            break;
                //        }
                //    }
                //}
                

                // ws.Application.Quit();

                Console.WriteLine("\n\nДля завершения работы нажмите л. кл.");
                Console.ReadLine();
            }
            catch (Exception)
            {

                throw;
            }
        }
        */
        #endregion

        // работает из под КАД
        static void Metod3TestDial()
        {
            var ed = new DataExcel();
            ed.TestDial();

        }
    }
}
