using System;
using System.Collections.Generic;
using System.Linq;
//using System.Data;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.IO;
using System.Windows.Forms;

//#error version

namespace ExcelData.Class
{
    public class DataExcel
    {

        private string fileExcelName;
        private string sheetExcelName;

        public DataExcel (string fileExcelName, string sheetExcelName)
        {
            this.fileExcelName = fileExcelName;
            this.sheetExcelName = sheetExcelName;
        }

        public DataExcel() { }

        public string [,] GetDataExel()
        {
            try
            {
#if !DEBUG
                OpenFileDialog dialog = new OpenFileDialog();

                DialogResult res = dialog.ShowDialog();
                if (res == DialogResult.OK)
                {
                    ExcelPackage excelFile = new ExcelPackage(
                        new FileInfo(dialog.FileName));

                    // спросим имя листа
                    using (Prompt prompt = new Prompt("Название листа EXCEL", "ВВЕДИТЕ ДАННЫЕ"))
                    {
                        sheetExcelName = prompt.Result;
                    }
#else

                //fileExcelName= Const.FileXlsName;
                if (File.Exists(fileExcelName))
                {

                ExcelPackage excelFile = new ExcelPackage(
                        new FileInfo(fileExcelName));
#endif



                    ExcelWorksheet worksheet =
                            excelFile.Workbook.
                            Worksheets[sheetExcelName];

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

                    return excelTable; //worksheet.Name.ToString();
                }
                else
                {
#if !DEBUG
                    throw new Exception("Файл не выбран!");
#else
                   return null;
#endif
                }
            }

            catch (Exception e)
            {
#if DEBUG
                throw;
#endif
                return null;
            }
        }

        public void TestDial ()
        {
            OpenFileDialog dialog = new OpenFileDialog();
            DialogResult res = dialog.ShowDialog();
        }
    }
}
