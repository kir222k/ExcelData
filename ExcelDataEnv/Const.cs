/* Кирилл Уваров 2022г. 10 февраля. u.k.send@gmail.com. +79062644029
*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelData

{
    public static class Const
    {
        //public static string FileXlsName = "u:\\dev\\ExcelConnect\\_Test\\ВРУ_ТРН.xlsx";
        public static string FileXlsName = "u:\\dev\\ExcelConnect\\_Test\\ВРУ_ТРН_2.xlsx";
        //public static string FileXlsName = "u:\\dev\\ExcelConnect\\_Test\\ArraTest2.xlsx";

        public static string LogFileName = "u:\\dev\\ExcelConnect\\_Test\\log.log";
        public static string LogFileTable = "u:\\dev\\ExcelConnect\\_Test\\table.log";

        public static string ExcelWorksheet = "Расчет";

        //public static List<string> ListGroupAttrsTest = new List<string> {
        //    "N.АПП1",
        //    "НАИМЕНОВАНИЕ.НАГРУЗКИ",
        //    "УЧАСТОК"};

        public static string NullTextReplace = "<emp>";

        public static string ExcelTextCellAsBblock = "[Блок]";
        public static string ExcelTextCellAsAttribute = "[Атрибут]";

        public static string BlockAttrApparatQF = "N.АПП1";
        public static string BlockAttrApparatSect = "УЧАСТОК";

        public static string XDataKeyExcelFilePath = "samexceldatapath";
        public static string XDataKeyExcelSheetName = "samexceldatasheet";


    }

    public static class Messg
    {
        /// <summary>
        /// Файл не существует! Требуется подключить файл Excel.
        /// </summary>
        public static string NotFile = "Файл не существует! Требуется подключить файл Excel."; // Messg.NotFile


        /// <summary>
        /// Файл не существует! Требуется подключить файл Excel.
        /// </summary>
        public static string NeedConnectExcelFile = "Для загрузки данных требуется выбрать файл Excel."; // Messg.NotFile

        // "Для связи данных требуется выбрать файл."



        /// <summary>
        /// Для связи данных требуется задать имя листа в книге Excel.
        /// </summary>
        public static string NeedSheetNameToConnect = "Для связи данных требуется задать имя листа в книге Excel."; // Messg.NotFileToConnect

        /// <summary>
        /// Такого листа не существует!
        /// </summary>
        public static string NotExcelSheet = "Такого листа не существует!"; // Messg.NotExcelSheet




        public static string AfterCancelDialogFile = "Загрузка данных не будет выполнена."; // Messg.OkExcelFileSheetConnect
        public static string OkExcelFileSheetConnect = "Ok."; // Messg.OkExcelFileSheetConnect

    }

}
