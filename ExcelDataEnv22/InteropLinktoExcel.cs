// http://wladm.narod.ru/C_Sharp/comexcel.html#0

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;


namespace ExcelData
{
    public  class InteropLinktoExcel
    {
    
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
