using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelData

{
    public static class Const
    {
        public static string FileXlsName = "u:\\dev\\ExcelConnect\\_Test\\ВРУ_ТРН.xlsx";
        //public static string FileXlsName = "u:\\dev\\ExcelConnect\\_Test\\ArraTest2.xlsx";

        public static string LogFileName = "u:\\dev\\ExcelConnect\\_Test\\log.log";

        public static string ExcelWorksheet = "Расчет";

        // 
        public static List<string> ListGroupAttrsTest = new List<string> {
            "N.АПП1",
            "НАИМЕНОВАНИЕ.НАГРУЗКИ",
            "УЧАСТОК"};

        public static string NullTextReplace = "<emp>";

        public static string ExcelTextCellAsBblock = "[Блок]";
        public static string ExcelTextCellAsAttribut = "[Атрибут]";
    }

  

}
