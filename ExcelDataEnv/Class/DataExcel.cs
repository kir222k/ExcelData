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
//#error version

namespace ExcelData.Class
{
    public class DataExcel
    {

        private string fileExcelName="";
        private string sheetExcelName="";

        public DataExcel (string fileExcelName, string sheetExcelName)
        {
            this.fileExcelName = fileExcelName;
            this.sheetExcelName = sheetExcelName;
        }

        public DataExcel() { }



        //public string[,] GetDataExel()
        public ArrayWithComments GetDataExel()
        {
            try
            {
                // Если имя файла и имя листа не равны "", то
                if ((fileExcelName != "") && (sheetExcelName != ""))
                {
                    // проверим файл на сущ.
                    if (File.Exists(fileExcelName))
                    {
                        // Создадим объект для работы с Excel
                        ExcelPackage excelFile = new ExcelPackage(new FileInfo(fileExcelName));

                        // проверим есть ли в файле лист с названием sheetExcelName
                        // считаем, что проверили:
                        // создаем объект для работы с листом 
                        ExcelWorksheet worksheet = excelFile.Workbook.Worksheets[sheetExcelName];

                        return new ArrayWithComments { Array = GetDataExelToArray(worksheet), Comments = "ok.." };
                    }
                    else
                    {
                        // говорим, что нет такого файла и просим переподключить
                        // throw new Exception("Файл не существует! Требуется переподключение.");
                        //MessageBox.Show("Файл не существует! Требуется подключить файл Excel.");
                        //return null;
                        return new ArrayWithComments {Array= null, Comments = "Файл не существует! Требуется подключить файл Excel." };
                    }

                // если файл или лист не заданы, то
                }
                else
                {
#if !DEBUG
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

                        // спросим имя листа
                        using (Prompt prompt = new Prompt("Название листа EXCEL", "ВВЕДИТЕ ДАННЫЕ"))
                        {
                            
                            sheetExcelName = prompt.Result;

                        }

                        //проверим на ""
                        if (sheetExcelName != string.Empty)
                        {
                            // проверим есть ли в файле лист с названием sheetExcelName
                            // считаем, что проверили:
                            // создаем объект для работы с листом 
                            ExcelWorksheet worksheet = excelFile.Workbook.Worksheets[sheetExcelName];
                            return new ArrayWithComments { Array = GetDataExelToArray(worksheet), Comments = "ok.." };
                        }
                        else
                        {
                            //MessageBox.Show("Для загрузки данных требуется задать имя листа в книге Excel.");
                            //return null;
                            return new ArrayWithComments { Array = null, Comments = "Для загрузки данных требуется задать имя листа в книге Excel." };
                        }

                    }


                        // иначе говорим, что файл не выбран и прерываем
                    else
                    {
                        //throw new Exception("Файл не выбран!");
                        //MessageBox.Show("Для загрузки данных требуется выбрать файл.");
                        //return null;
                        return new ArrayWithComments { Array = null, Comments = "Для загрузки данных требуется выбрать файл." };
                    }
#else
                    return new ArrayWithComments { Array = null, Comments = "Не заданы имя файла или листа Excel!" };
#endif
                }
            }

            catch (Exception e)
            {
                //return null;
                return new ArrayWithComments { Array = null, Comments = "Неизвестная ошибка." };
            }
        }


       private string[,] GetDataExelToArray(ExcelWorksheet worksheet)
        {
                int totalRows = worksheet.Dimension.End.Row;
                int totalColumns =  worksheet.Dimension.End.Column;

                string [,] excelTable = new string[totalRows, totalColumns];

                for (int i = 0; i < totalRows - 1; i++)
                {
                    for (int j = 0; j < totalColumns - 1; j++)
                    {
                        if (worksheet.Cells[i+1, j+1].Value != null)
                        {
                            excelTable[i, j] = Convert.ToString(worksheet.Cells[i+1, j+1].Value);
                        }
                        else
                        {
                            excelTable[i, j] = Const.NullTextReplace;
                        }
                    }
                }

                return excelTable; 

        }




        public void TestDial ()
        {
            OpenFileDialog dialog = new OpenFileDialog();
            DialogResult res = dialog.ShowDialog();
        }
    }
}
