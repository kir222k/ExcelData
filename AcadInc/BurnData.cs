using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using System.Reflection;
using Autodesk.Windows;
using acadApp = Autodesk.AutoCAD.ApplicationServices.Application;
using ExcelData;
using ExcelData.Class;
using ExcelData.Model;
using System.Windows.Forms;


[assembly: CommandClass(typeof(AcadInc.DataFrom))]


namespace AcadInc
{
    /// <summary>
    /// Загрузка даных из Excel.
    /// </summary>
    public static class DataFrom
    {

        /// <summary>
        /// Основной метод загрузки данных.
        /// </summary>
        [CommandMethod("BurnDataFromExcel")]
        public static void BurnData()
        {
            // проверка - если еще нет файла EXCEL (прочитать путь-строку из расш. данных DWG файла )
            BurnDataDial();

            // Если есть в расш. даных dwg файла, то
            // проверить на сущ., если есть такой на диске (путь не сломан)
            //BurnDataSavedPath();
            // если нет, то выдать сообщ. об ошибке, и запустить диалог выбора файла
            // BurnDataDial();


        }

        public static void BurnDataDial()
        {
            var AcSd = new AcadSendMess();
            DataExcel DE = new DataExcel();

            // получим путь к файлу и имя листа, кот. были заданы нами и кот. нужно запомнить в расш. данных
            var tuple = BurnDataBased(DE);

            // выведем их
            // AcSd.SendStringDebugStars(tuple.file + "\n" + tuple.sheet);
            AcSd.SendStringDebugStars(new List<string>
            {
                tuple.file,
                tuple.sheet
            });
        }

        public static void BurnDataSavedPath()
        {
            // DataExcel ED1 = new DataExcel(fileExcelName: <путь из расш.данных файла DWG>, sheetExcelName: "Расчет");
            
        }

        private static (string file, string sheet) BurnDataBased(DataExcel DE)
        {

            var AcSd = new AcadSendMess();
            var ArrComm = new ArrayWithComments();
            ArrComm= DE.GetDataExel();

            string[,] strTable = ArrComm.Array;

            if (strTable != null)
            {
                int rows = strTable.GetUpperBound(0) + 1;    // количество строк
                int columns = strTable.GetUpperBound(1) + 1; // количество столбцов

                // Console.WriteLine($"Строк={rows}  Столбцов={columns}");
                // var str = $"Строк={rows}  Столбцов={columns}";

                AcSd.SendStringDebugStars(ArrComm.Comments);

                return (DE.FileExcelName,DE.SheetExcelName);

            }
            else
            {
                //var AcSd = new AcadSendMess();
                //AcSd.SendStringDebugStars(ArrComm.Comments);
                MessageBox.Show(ArrComm.Comments);
                return (string.Empty, string.Empty);
            }

            var PP = new PullPushData(strTable);
            AcSd.SendStringDebug(PP.ToString());

        }

    }
}
