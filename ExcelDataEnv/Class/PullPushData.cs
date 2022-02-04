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
        public List<BlockData> GetListBlockDataToPush()
        {
            //throw new NotImplementedException();

            return new List<BlockData>
            {
                new BlockData
                {BlockName="Line_load3", ListAttributes=
                    new List<AttrData>
                    {
                        new AttrData {AttributeTag="НАИМЕНОВАНИЕ.НАГРУЗКИ", 
                            AttributeValue="Квартирный стояк N1.1 (секция 3)"},
                        new AttrData {AttributeTag="N.АПП1", AttributeValue="QF1"},
                        new AttrData {AttributeTag="УЧАСТОК", AttributeValue="1"}
                    }
                },
                new BlockData
                {BlockName="Line_load3", ListAttributes=
                    new List<AttrData>
                    {
                        new AttrData {AttributeTag="НАИМЕНОВАНИЕ.НАГРУЗКИ", 
                            AttributeValue="Квартирный стояк N1.2 (секция 3)"},
                        new AttrData {AttributeTag="N.АПП1", AttributeValue="QF2"},
                        new AttrData {AttributeTag="УЧАСТОК", AttributeValue="1"}
                    }
                }

            };
        }

        public override string ToString()
        {
            string  str="";

            foreach (var BlData in GetListBlockDataToPush())
            {
                str += "\n\nБлок: " + BlData.BlockName;
                str += "\nАтрибуты=>";
                int it = 1;
                foreach (var attrs in BlData.ListAttributes)
                {
                    str += "\n" + it.ToString() + ":";
                    str +=  "\nAttributeTag: "   + attrs.AttributeTag + 
                            "\nAttributeValue: "  + attrs.AttributeValue;
                    it++;
                }
            }

            return str;

        }


    }
}
