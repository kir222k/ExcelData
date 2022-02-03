using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelData.Model;

namespace ExcelData
{
    public class PullPushData: IPullPushData
    {
        // private string fileExcelName;
        // private string sheetExcelName;
        private string[,] strTable;

        // Конструктор - создание листа Excel из файла библиотекой EPPlus
        public PullPushData (string[,] strTable)  //, string fileExcelName, string sheetExcelName)
        {
            // this.fileExcelName = fileExcelName;
            // this.sheetExcelName = sheetExcelName;
            this.strTable = strTable;
        }

        public PullPushData() { }

        // Блок.
        public ExcelRangeText GetExcelCellBlockText()
        {
            //throw new NotImplementedException();

            return new ExcelRangeText { TextValue = "Тута", ColumnCell = 1, RowCell = 2 };
        }

        // Атрибут.
        public ExcelRangeText GetExcelCellAttributeText()
        {
            throw new NotImplementedException();
        }


        // Список блоков.
        public List<ExcelRangeText> GetExcelRangeBlock()
        {
            throw new NotImplementedException();
        }

        // Список атр.
        public List<ExcelRangeText> GetExcelRangeAttribute()
        {
            throw new NotImplementedException();
        }

        // Список данных для выгрузки.
        public  List<BlockData> GetListBlockDataToPush()
        {
            throw new NotImplementedException();
        }

        public override string ToString()
        {
            return "PullPushData!!";

        }


    }
}
