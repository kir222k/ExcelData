using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using OfficeOpenXml;
using System.IO;
using System.Windows.Forms;
using ExcelData.Model;
using ExcelData;


namespace ExcelData.Class
{
    public  static class DataCheck
    {
        public static (bool isFile, bool isSheet) IsExelSheetExist(string fileExcelName, string sheetExcelName)
        {
            // DataCheck.IsExelSheetExist (fileExcelName,sheetExcelName)
            bool isExistFile = false;
            bool isExistWorksheet = false;
            // проверим файл на сущ.
            if (File.Exists(fileExcelName))
            {
                isExistFile = true;
                // Создадим объект для работы с Excel
                ExcelPackage excelFile = new ExcelPackage(new FileInfo(fileExcelName));

                // проверим есть ли в файле лист с названием sheetExcelName
                var WS = excelFile.Workbook.Worksheets;
                foreach (var item in WS)
                {
                    if (item.Name == sheetExcelName)
                        isExistWorksheet = true;
                }
                excelFile.Dispose();
            }
            return (isExistFile, isExistWorksheet);
        }

    }
}
