// http://wladm.narod.ru/C_Sharp/comexcel.html#0

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace ExcelData
{
    public  class ExcelBook
    {
        //static Excel.Application excelapp = new Excel.Application() { Visible = false };

        // Диалог. окно - указать файл
        // запомнить имя файла в переменную
        // Пока это - Const.FileXlsName

        // Метод считывания из опред. книги, листа опр. данных
        // Реализация - получить список типа List<Group> из данных листа
        /*
        public static List<Group> GetListDataBoard (string nameBlockDef, List<string> nameAttributes)
        {
            // атрибут
            AttrData attr;
            // Группа с атрибутами
            Group gr;
            // Список групп с атрибутами
            List<Group> listGr=new List<Group>();
            
            // 2. Получаем ссылку на лист 
            Excel.Worksheet excelworksheet = GetExcelSheet(Const.FileXlsName, Const.ExcelWorksheet);

            // 3. Получим массив 50*50 значений  от А1, поищем там ячейку с текстом "Блок"
            int x1 = 200; int y1 = 150; string range1 = "A1";
                        var ArrForBlock = GetArrayBasedCell(ref excelworksheet, x1, y1, range1);
            for (int i = 0; i < x1; i++)
            {
                for (int j = 0; j < y1; j++)
                {
                    // Console.Write(Arr[i, j].ToString() + " ");
                    if (ArrForBlock[i,j]=="Блок")
                    {
                        // кортеж с данными (строка, столбец) , где сидит слово "Блок"
                        var rangeBlock = (i, j); // столбец, в кот. будут имена блоков найден
                        break;
                    }
                }
            }

            // получим данные (i, j) где сидит слово "Атрибут:"
            int x2 = 50; int y2 = 50; string range2 = "A1";
            var ArrForAttrib = GetArrayBasedCell(ref excelworksheet, x2, y2, range2);
            for (int i = 0; i < x2; i++)
            {
                for (int j = 0; j < y2; j++)
                {
                    // Console.Write(Arr[i, j].ToString() + " ");
                    if (ArrForBlock[i, j] == "Атрибут:")
                    {
                        // кортеж с данными (строка, столбец) , где сидит слово "Атрибут:"
                        var rangeAttrib = (i, j); // строка, в кот. будут имена атрибутов найден
                        break;
                    }
                }
            }
            




            return listGr ?? null;
        }


        //public static string  ExcelBookTest()
        //{
        //    //Получаем ссылку на лист 
        //    Excel.Worksheet excelworksheet = GetExcelSheet(Const.FileXlsName, Const.ExcelWorksheet);
        //    //Выбираем ячейку для вывода В4
        //    var excelcells = excelworksheet.get_Range("B4", Type.Missing);

        //    // конв. в строку
        //    var sStr = Convert.ToString(excelcells.Value2);

        //    // закроем лист
        //    //excelappworkbook.Close();
        //    //excelworksheet.Application.ThisWorkbook.Close();
        //    try
        //    {
        //        //excelworksheet.Application.ThisWorkbook.Close();
        //        // закроем экз. Excel 
        //        //excelapp.Quit();
        //        excelworksheet.Application.Quit();

        //    }
        //    catch (Exception)
        //    {
        //        Console.WriteLine("Ошибка закрытия . Проверить диспетчер");
        //        //throw;
        //    }


        //    return sStr;
        //}

        ///// <summary>
        ///// Возращает лист с данными.
        ///// </summary>
        ///// <param name="path">Путь к файлу Excel</param>
        ///// <param name="sheetname">Назв. листа</param>
        ///// <returns></returns>
        //public static  Excel.Worksheet GetExcelSheet( string path, string sheetname)
        //{
        //    // получаем объект Excel
        //    //excelapp = new Excel.Application();
            
        //    //excelapp.Visible = false; // если нужно показать - true
        //    // окрываем сущ. файл
        //    // https://docs.microsoft.com/ru-ru/office/vba/api/excel.workbooks.open
        //    /*
        //     Workbook_object.Open(
        //         FileName,         //Имя открываемого файла файла
        //         UpdateLinks,      //Способ обновления ссылок в файле
        //         ReadOnly,         //При значении true открытие только для чтения 
        //         Format,           //Определение формата символа разделителя
        //         Password,         //Пароль доступа к файлу до 15 символов
        //         WriteResPassword, //Пароль на сохранение файла
        //         IgnoreReadOnlyRecommended, //При значении true отключается вывод 
        //                                    //запроса на работу без внесения изменений
        //         Origin,           //Тип текстового файла 
        //         Delimiter,        //Разделитель при Format = 6
        //         Editable,         //Используется только для надстроек Excel 4.0
        //         Notify,           //При значении true имя файла добавляется в 
        //                           //список нотификации файлов
        //         Converter,        //Используется для передачи индекса конвертера файла
        //                           //используемого для открытия файла    
        //         AddToMRU          //При true имя файла добавляется в список 
        //                           //открытых файлов
        //                         ) 
        //    */
        //    var excelappworkbook = excelapp.Workbooks.Open(Filename: path, ReadOnly: true);
        //    // Получаем его листы
        //    var excelsheets = excelappworkbook.Worksheets;
        //    // Получаем ссылку на лист
        //    var excelworksheet = (Excel.Worksheet)excelsheets.get_Item(sheetname);

        //    // закроем экз. Excel 
        //    excelapp.Quit();
        //    excelworksheet.Application.Quit();

        //    return excelworksheet;
        //}

        /// <summary>
        /// Возращает массив из  х строк у столбцов от ячейки
        /// </summary>
        /// <param name="x">строк [x,y]</param>
        /// <param name="y">столбцов [x,y]</param>
        /// <param name="range">назв. ячейки от которой отсчитывать строки вправо, стролбцы вниз</param>
        /// <returns></returns>
        public string [,] GetArrayBasedCell (string pathFile, string sheetName, int x, int y, string rangeName)
        {
            try
            {

            Excel.Application excelapp = new Excel.Application() { Visible = false };
            var excelappworkbook = excelapp.Workbooks.Open(Filename: pathFile, UpdateLinks: false, ReadOnly: true);



            // Получаем его листы
            var excelsheets = excelappworkbook.Worksheets;
            // Получаем ссылку на лист
            var excelworksheet = (Excel.Worksheet)excelsheets.get_Item(sheetName);

            // переменнная массива
            string[,] ArrayData= new string[x, y];

            int rangeX= (int)excelworksheet.get_Range(rangeName).Row; 
            int rangeY= (int)excelworksheet.get_Range(rangeName).Column ;
            string str = "";

            Excel.Range excelcells;

            // возьмем строки
            for (int i = 0; i < x; i++)
            {
                // столбцы
                for (int j = 0; j < y; j++)
                {
                    excelcells = (Excel.Range)excelworksheet.Cells[rangeX + i, rangeY + j];

                    if (excelcells.Value2 != null)
                    {
                        str = Convert.ToString(excelcells.Value2);
                    }
                    else
                    {
                        str = "NULL";
                    }
                     
                    ArrayData[i, j] = str;

                }

            }

            excelappworkbook.Close();
            // закроем экз. Excel 
            excelapp.Quit();
           

            return ArrayData ?? null;

            }
            catch (Exception)
            {
                return null;
             
            }
        }


    }
}
