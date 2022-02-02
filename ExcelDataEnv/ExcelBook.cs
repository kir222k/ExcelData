// http://wladm.narod.ru/C_Sharp/comexcel.html#0



using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace ExcelData
{
    public static class ExcelBook
    {


        // Диалог. окно - указать файл
        // запомнить имя файла в переменную
        // Пока это - Const.FileXlsName

        // Метод считывания из опред. книги, листа опр. данных
        // Реализация - получить список типа List<Group> из данных листа
        public static List<Group> GetListDataBoard (string nameBlockDef, List<string> nameAttributes)
        {
             Excel.Application excelapp;
             Excel.Window excelWindow;
            excelapp = new Excel.Application();
            excelapp.Visible = true;

            // атрибут
            AttrData attr;

            // Группа с атрибутами
            Group gr;

            // Список групп с атрибутами
            List<Group> listGr=new List<Group>();

            /*
             * 1. Получить объект-книгу Excel
             * 2. Получить объект-лист Excel
             * 3. ЦИКЛ слева на право по столбцам 
             *       - ЦИКЛ->поискать в первых 50 строках ячейку с текстом "Блок" - запомнить этот слолбец! 
             *       тут будут имена блоков
             * 4. Находим ячейку, в кот. текст "Атрибут:" 
             * 5. Начать проходить по слолбцу, кот. назначен как столбец имен блоков 
             *      - ЦИКЛ-> находим имя блока,  gr.Name = имя блока, далее - 
             *      идем по строке и вносим пары в listGr пары <атрибут.значение> типа GroupData
             */

            return listGr ?? null;
        }


        public static string  ExcelBookTest()
        {
            // получаем объект Excel
            Excel.Application excelapp = new Excel.Application();
            excelapp.Visible = false; // если нужно показать - true

            //Excel.Workbooks excelappworkbooks = excelapp.Workbooks;

            // окрываем сущ. файл
            // https://docs.microsoft.com/ru-ru/office/vba/api/excel.workbooks.open
            //var excelappworkbook =excelapp.Workbooks.Open(Const.FileXlsName,
            //  Type.Missing, true , Type.Missing, Type.Missing,
            //  Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            //  Type.Missing, Type.Missing, Type.Missing, Type.Missing,
            //  Type.Missing, Type.Missing);
            
            var excelappworkbook = excelapp.Workbooks.Open(Filename: Const.FileXlsName, ReadOnly: true);

            // Получаем его листы
            var excelsheets = excelappworkbook.Worksheets;

            //Получаем ссылку на лист 
            var excelworksheet = (Excel.Worksheet)excelsheets.get_Item(Const.ExcelWorksheet); // в моем случае "Расчет"
            //Выбираем ячейку для вывода В4
            var excelcells = excelworksheet.get_Range("B4", Type.Missing);

            // конв. в строку
            var sStr = Convert.ToString(excelcells.Value2);

            // закроем лист
            excelappworkbook.Close();
            // закроем экз. Excel 
            excelapp.Quit();
            
            return sStr;
        }

        public static Excel.Worksheet GetExcelSheet(string path, string sheetname)
        {
            return null;
        }

        /*
         Workbook_object.Open(
             FileName,         //Имя открываемого файла файла
             UpdateLinks,      //Способ обновления ссылок в файле
             ReadOnly,         //При значении true открытие только для чтения 
             Format,           //Определение формата символа разделителя
             Password,         //Пароль доступа к файлу до 15 символов
             WriteResPassword, //Пароль на сохранение файла
             IgnoreReadOnlyRecommended, //При значении true отключается вывод 
                                        //запроса на работу без внесения изменений
             Origin,           //Тип текстового файла 
             Delimiter,        //Разделитель при Format = 6
             Editable,         //Используется только для надстроек Excel 4.0
             Notify,           //При значении true имя файла добавляется в 
                               //список нотификации файлов
             Converter,        //Используется для передачи индекса конвертера файла
                               //используемого для открытия файла    
             AddToMRU          //При true имя файла добавляется в список 
                               //открытых файлов
) 
        */


    }
}
