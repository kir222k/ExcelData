/* Кирилл Уваров 2022г. 10 февраля. u.k.send@gmail.com. +79062644029
*/

// Чтобы увидеть реакцию системы на ошибки =  IsThrow = true
// Встречается там, где выявлены и пролечены исключения,
// если нужно быстро отключить эти таблетки и увидеть ,как именно происходят вылеты
#define IsThrow

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
            TypedValue[] valsPath = ExtData.ReadAndGetExtDataModel(Const.XDataKeyExcelFilePath).valueX;
            if (valsPath != null)
            {
                if  (valsPath.Count() == 1 ) 
                {
                    string pathFile = valsPath[valsPath.Count() - 1].Value.ToString();
                    if (pathFile != string.Empty)
                    {

                        TypedValue[] valsSheet = ExtData.ReadAndGetExtDataModel(Const.XDataKeyExcelSheetName).valueX;
                        if (valsSheet != null)
                        {
                            if (valsSheet.Count() == 1)
                            {
                                string sheetFile = valsSheet[valsSheet.Count() - 1].Value.ToString();
                                if (sheetFile != string.Empty)
                                {
                                    // чтобы не вылетало при попытке загрузки файла, кот.нет
                                    // чтобы посмотреть вылет => !IsThrow
#if IsThrow
                                    if (System.IO.File.Exists(pathFile))
                                        if (DataCheck.IsExelSheetExist(pathFile, sheetFile).isSheet) // или листа кот.нет
#endif
                                            BurnDataSavedPath(pathFile, sheetFile);
#if IsThrow
                                        else
                                            MessageBox.Show($"Лист \"{sheetFile}\" в связанном файле \n{pathFile}\n поврежден или отстутствует"); 
                                    else
                                        MessageBox.Show($"Связанный файл \n\"{pathFile}\"\n поврежден или отстутствует");
#endif

                                }
                            }
                        }
                    }
                }
            }
            else
            {
                // запустить диалог выбора файла
                BurnDataDial();

            }
        }


        [CommandMethod("BurnDataFromExcelReplace")]
        public static void BurnDataDial()
        {
            var AcSd = new AcadSendMess();
            DataExcel DE = new DataExcel();

            // получим путь к файлу и имя листа, кот. были заданы нами и кот. нужно запомнить в расш. данных
            var tuple = BurnDataBased(DE);

            // Если имя файла и листа не пустые
            if (
                //(tuple.file != string.Empty) && 
                //(tuple.sheet != string.Empty) &&
                //(tuple.blockDatas  != null) // если не  добавить проверку tuple.blockDatas на null, будет вылет при отмене диалог. окон.? непонятно пока!
                (tuple.file != string.Empty) &&
                (tuple.sheet != string.Empty) 

               )
            {
                // заберем путь м имя листа
                (string, string) dataToExtData = (tuple.file, tuple.sheet);
                // Отправим на запись в расш. данные
                ExtData.WriteToExtDataExcelFileInfo(dataToExtData);

                // а  tuple.blockDatas Передадим в класс, кот. занесет данные в атрибуьы блока
                BlockData.BlockRefModifity(tuple.blockDatas);
            }
            else 
            {
                // AcSd.SendStringDebugStars("Data empty");
            }
        }

        public static void BurnDataSavedPath(string fileExcelName, string sheetExcelName)
        {
            // создадим экз. класса для работы с данными из Excel
            DataExcel DE = new DataExcel(fileExcelName, sheetExcelName);

            // получим данные
            var tuple = BurnDataBased(DE);

            // а  tuple.blockDatas Передадим в класс, кот. занесет данные в атрибуты блока
            BlockData.BlockRefModifity(tuple.blockDatas);

        }

        private static (string file, string sheet, List<ExcelData.Model.BlockData> blockDatas) 
            BurnDataBased(DataExcel DE)
        {

            var AcSd = new AcadSendMess();
            var ArrComm = new ArrayWithComments();
            ArrComm = DE.GetDataExel();

            string[,] strTable = ArrComm.Array;

            if (strTable != null)
            {
                PullPushData PP = new PullPushData(strTable);
                return (DE.FileExcelName,DE.SheetExcelName, PP.GetListBlockDataToPush());
            }
            else
            {
                //var AcSd = new AcadSendMess();
                //AcSd.SendStringDebugStars(ArrComm.Comments);
                if (ArrComm.Comments != Messg.AfterCancelDialogFile) // пропускаем сообщение при закрытии диал. окна по кhестику или ESCAPE (юзер и так понимает, что отказывается от продолжения)
                {

                    MessageBox.Show(ArrComm.Comments);
                }
                return (string.Empty, string.Empty, null);
            }
        }
    }
}
