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
        //public void DataToAcad()
        //{
        //    DataExcel ED = new DataExcel();
        //    string[,] strTable = ED.GetDataExel();
        //}

        private string fileExcelName;
        private string sheetExcelName;


        // Конструктор - создание листа Excel из файла библиотекой EPPlus
        public PullPushData (string fileExcelName, string sheetExcelName)
        {
            this.fileExcelName = fileExcelName;
            this.sheetExcelName = sheetExcelName;
        }

        // Атрибут.
        ExcelRangeText IPullPushData.GetExcelCellAttributeText()
        {
            throw new NotImplementedException();
        }

        // Блок.
        ExcelRangeText IPullPushData.GetExcelCellBlockText()
        {
            throw new NotImplementedException();
        }

        // Список блоков.
        List<ExcelRangeText> IPullPushData.GetExcelRangeBlock()
        {
            throw new NotImplementedException();
        }

        // Список атр.
        List<ExcelRangeText> IPullPushData.GetExcelRangeAttribute()
        {
            throw new NotImplementedException();
        }

        // Список данных для выгрузки.
        List<BlockData> IPullPushData.GetListBlockDataToPush()
        {
            throw new NotImplementedException();
        }
    }
}
