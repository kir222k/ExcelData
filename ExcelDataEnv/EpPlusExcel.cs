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

namespace ExcelData
{
    public class EpPlusExcel
    {
        private string[,] exceTable;
        private int totalRows = 0;
        private int totalColumns = 0;

        public string GetDataExel()
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
#else
                if (File.Exists(Const.FileXlsName))
                {

                
                ExcelPackage excelFile = new ExcelPackage(
                        new FileInfo(Const.FileXlsName));
#endif

                    ExcelWorksheet worksheet =
                            excelFile.Workbook.
                            Worksheets[Const.ExcelWorksheet];

                    totalRows = worksheet.Dimension.End.Row;
                    totalColumns = worksheet.Dimension.End.Column;

                    exceTable = new string[totalRows, totalColumns];

                    for (int rowIndex = 1; rowIndex <= totalRows; rowIndex++)
                    {
                        IEnumerable<string>  row =
                            worksheet.Cells[rowIndex, 1, rowIndex, totalColumns].
                            Select(c => c.Value == null ? string.Empty : c.Value.ToString());

                        List<string> list = row.ToList<string>();

                        for (int i = 0; i < list.Count; i++)
                        {
                            exceTable[rowIndex - 1, i] =
                                Convert.ToString(list[i].Replace(".", ","));
                        }


                    }


                    return null; //worksheet.Name.ToString();
                }

                
                else
                {
                    //throw new Exception("Файл не выбран!");
                    return "Файл не выбран!";
                }
            }
            catch (Exception e)
            {
                return e.Message.ToString();
                //throw;
            }
        }

        public void TestDial ()
        {
            OpenFileDialog dialog = new OpenFileDialog();
            DialogResult res = dialog.ShowDialog();
        }
    }
}
