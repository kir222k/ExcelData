/* Кирилл Уваров 2022г. 10 февраля. u.k.send@gmail.com. +79062644029
*/

#define IsVariant

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
    public class DataExcel
    {
        private string fileExcelName="";
        private string sheetExcelName="";

        public string FileExcelName { get => fileExcelName; }
        public string SheetExcelName { get => sheetExcelName; }

        public DataExcel (string fileExcelName, string sheetExcelName)
        {
            this.fileExcelName = fileExcelName;
            this.sheetExcelName = sheetExcelName;
        }

        public DataExcel() { }

        /// <summary>
        /// Возращает дынные из листа книги Excel по заданному пути.
        /// </summary>
        /// <returns>2мерный массив с комментарием, т.е. если получаем  null, то из описания понятно, почему</returns>
        public ArrayWithComments GetDataExel()
        {
            try
            {
                // Если имя файла и имя листа не равны "", то
                if ((fileExcelName != string.Empty) && (sheetExcelName != string.Empty))
                {
                    if (DataCheck.IsExcelSheetExist(fileExcelName, sheetExcelName).isFile)
                    {
                        ExcelPackage excelFile = new ExcelPackage(new FileInfo(fileExcelName));
                        ExcelWorksheet worksheet = excelFile.Workbook.Worksheets[sheetExcelName];

                        if (DataCheck.IsExcelSheetExist(fileExcelName, sheetExcelName).isSheet)
                            return new ArrayWithComments { Array = GetDataExelToArray(worksheet), Comments = Messg.OkExcelFileSheetConnect };
                        else
                            return new ArrayWithComments { Array = null, Comments = Messg.NotExcelSheet };
                    }
                    else
                        return new ArrayWithComments { Array = null, Comments = Messg.NotFile };

 #region УДАЛИТЬ?
                    // проверим файл на сущ.
                    ////if (File.Exists(fileExcelName))
                    ////{
                    ////    // Создадим объект для работы с Excel
                    ////    ExcelPackage excelFile = new ExcelPackage(new FileInfo(fileExcelName));

                    ////    // проверим есть ли в файле лист с названием sheetExcelName
                    ////    bool isExistWorksheet = false;
                    ////    var WS = excelFile.Workbook.Worksheets;
                    ////    foreach (var item in WS)
                    ////    {
                    ////        if (item.Name == sheetExcelName)
                    ////            isExistWorksheet = true;
                    ////    }

                    ////    if (isExistWorksheet)
                    ////    {
                    ////        ExcelWorksheet worksheet = excelFile.Workbook.Worksheets[sheetExcelName];
                    ////        return new ArrayWithComments { Array = GetDataExelToArray(worksheet), Comments = "ok.." };
                    ////    }
                    ////    else
                    ////    {
                    ////        return new ArrayWithComments { Array = null, Comments = "Такого листа не существует!" };
                    ////    }

                    ////    // считаем, что проверили:
                    ////    // создаем объект для работы с листом 
                    ////}
                    ////else
                    ////{
                    ////    // говорим, что нет такого файла и просим переподключить
                    ////    // throw new Exception("Файл не существует! Требуется переподключение.");
                    ////    //MessageBox.Show("Файл не существует! Требуется подключить файл Excel.");
                    ////    //return null;
                    ////    return new ArrayWithComments {Array= null, Comments = "Файл не существует! Требуется подключить файл Excel." };
                    ////}
#endregion
                    // если файл или лист не заданы, то
                }
                else
                {
//#if !DEBUG
                    // откроем диалог
                    OpenFileDialog dialog = new OpenFileDialog();
                        // настроить, чтоб видны только *.xlsx
                        DialogResult res = dialog.ShowDialog();
                    // если  нажали ок после файла
                    if (res == DialogResult.OK)
                    {
                        // запомним имя файла
                        fileExcelName = dialog.FileName;
                        // Создадим объект для работы с Excel
                        ExcelPackage excelFile = new ExcelPackage(new FileInfo (fileExcelName));

                        // список листов книги
                        List<string> listSheet = new List<string>();
                        foreach (ExcelWorksheet ws in excelFile.Workbook.Worksheets)
                        {
                            listSheet.Add(ws.Name);
                        }

                        // спросим имя листа
                        using (Prompt prompt = new Prompt("Книга: " + fileExcelName , "Выберите лист в книге Excel", listSheet))
                        {
                            sheetExcelName = prompt.Result;
                        }

                        //проверим на ""
                        if (sheetExcelName != string.Empty)
                        {
                            // проверим есть ли в файле лист с названием sheetExcelName
                            bool isExistWorksheet = false;
                            // по всем листам книги:
                            var WS = excelFile.Workbook.Worksheets;
                            foreach (var item in WS)
                            {
                                if (item.Name == sheetExcelName) // если совпадение
                                    isExistWorksheet = true;
                            }
                            // Лист есть
                            if (isExistWorksheet)
                            {
                                // создаем объект для работы с листом 
                                ExcelWorksheet worksheet = excelFile.Workbook.Worksheets[sheetExcelName];
                                return new ArrayWithComments { Array = GetDataExelToArray(worksheet), Comments = Messg.OkExcelFileSheetConnect };
                            }
                            else // Листа нет
                            {
                                
                                return new ArrayWithComments { Array = null, Comments = Messg.NotExcelSheet };
                            }
                        }
                        else // если введена пустая строка
                        {
                            //MessageBox.Show("Для загрузки данных требуется задать имя листа в книге Excel.");
                            //return null;
                            return new ArrayWithComments { Array = null, Comments = Messg.NeedSheetNameToConnect };
                        }

                    }
                    // иначе говорим, что файл не выбран и прерываем
                    else
                    {
                        if (res == DialogResult.Cancel)
                        {
                            return new ArrayWithComments { Array = null, Comments = Messg.AfterCancelDialogFile };
                        }
                        //MessageBox.Show("Для загрузки данных требуется выбрать файл.");
                        //return null;
                        return new ArrayWithComments { Array = null, Comments = Messg.NeedConnectExcelFile };
                    }
//#else
//                    return new ArrayWithComments { Array = null, Comments = "Не заданы имя файла или листа Excel!" };
//#endif
                }
            }

            catch (Exception e)
            {
                //return null;
                return new ArrayWithComments { Array = null, Comments = $"Неизвестная ошибка.\n{e.Message.ToString()}" };
            }
        }


       internal protected string[,] GetDataExelToArray(ExcelWorksheet worksheet)
        {
                int totalRows = worksheet.Dimension.End.Row;
                int totalColumns =  worksheet.Dimension.End.Column;

                string [,] excelTable = new string[totalRows, totalColumns];

                for (int i = 0; i < totalRows - 1; i++)
                {
                    for (int j = 0; j < totalColumns - 1; j++)
                    {
                        ExcelRangeBase Cell = worksheet.Cells[i + 1, j + 1];
                        if (Cell.Value != null)
                        {

#if !IsVariant
                        // найдем числа количеством знаков после запятой, кот. больше, чем заданное. напр. 2.

                        // переведем значение ячейки в строку.
                        string strChek = Convert.ToString(Cell.Value);

                        if (
                            // (Cell.Style.Numberformat.Format.Contains("0.0")) || // Если формат числа .. уберем это из условия 
                            // проверяем , если это строка, кот. может быть преобразована в число -
                            // или все цифры или цифры с единств. разделителем "," или "."
                            (ValueChek.IsDigitStr(strChek)) &&
                            //(ValueChek.GetQuantOfPoint(strChek) > 0) // т.е. дробь
                            (ValueChek.GetQuantOfPoint(strChek) > Const.RoundForDouble) // проверяем, если число знаков  больше заданного.
                           )
                        {
                            // т.е. нашли строку, кот. может быть преобразована в число, и число знаков больше заданного
                            // для нашего случая заменим разделитель, если он "." на ","
                            //strChek = ValueChek.GetStringWithPointCorrect(strChek);
                            // преобразуем в число
                            double valueDouble = Convert.ToDouble(strChek);
                            // округлим
                            //string str = Convert.ToString(Math.Round(valueDouble, Const.RoundForDouble));
                            strChek = Convert.ToString(Math.Round(valueDouble, Const.RoundForDouble));
                            // добавим нули, если не хватает до нужного числа знаков после запятой
                            //str = ValueChek.GetAddZeroStr(str, Const.RoundForDouble); //- при данных усл. выполнение не имеет смысла
                            // запишем в массив
                            excelTable[i, j] = strChek;
                        }
                        else
                            excelTable[i, j] = Convert.ToString(Cell.Value);
#else
                        excelTable[i, j] = Convert.ToString(Cell.Text);
#endif
                    }
                    else
                        {
                            excelTable[i, j] = Const.NullTextReplace;
                        }
                    }
                }

                return excelTable; 

        }



    }
}
