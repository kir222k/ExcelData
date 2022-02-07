using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelData.Model;

namespace ExcelData.Class
{
    public class PullPushData: IPullPushData
    {
        private string[,] strTable;
        private ExcelRangeText excelCellBlockText;
        private ExcelRangeText excelCellAttributeText;

        #region КОНСТРУКТОРЫ
        /// <summary>
        /// Конструктор - передаем  лист Excel, созданный из  файла ТРН библиотекой EPPlus.
        /// </summary>
        /// <param name="strTable"> Лист Excel в виде 2мерн. массива.</param>
        public PullPushData (string[,] strTable)  //, string fileExcelName, string sheetExcelName)
        {
            // this.fileExcelName = fileExcelName;
            // this.sheetExcelName = sheetExcelName;
            this.strTable = strTable;
        }

        /// <summary>
        /// Конструктор без арг.
        /// </summary>
        public PullPushData() { }
        #endregion

        /*
        /// <summary>
        /// Ищем "[Блок]"
        /// </summary>
        /// <returns>Зн.ячейки с ее координатами.</returns>
        public ExcelRangeText GetExcelCellBlockText()
        {
            ExcelRangeText? eT = null;



            return (ExcelRangeText)eT;

            //Nullable<ExcelRangeText> eT=null;
            //throw new NotImplementedException();
            //return new ExcelRangeText { TextValue = "Тута", ColumnCell = 1, RowCell = 2 };
        }

        /// <summary>
        /// Ищем "[Атрибут]"
        /// </summary>
        /// <returns>Зн.ячейки с ее координатами.</returns>
        public ExcelRangeText GetExcelCellAttributeText()
        {
            throw new NotImplementedException();
        }
        */

        // Список блоков.
        public List<ExcelRangeText> GetExcelRangeBlock()
        {
            var listBlocks = new List<ExcelRangeText>();
            // Ищем "[Блок]" и координаты ячеки с этим текстом.
            //if (excelCellBlockText == null)
            excelCellBlockText = SearchValueInArray.GetCellCoordinatesInArray(Const.ExcelTextCellAsBblock, strTable);

            int rows = strTable.GetUpperBound(0) + 1;    // количество строк
            //int columns = strTable.GetUpperBound(1) + 1; // количество столбцов
            // берем массив и ищем в слолбце excelCellBlockText.ColumnCell
            int j = excelCellBlockText.ColumnCell;
            for (int i = 0; i < rows - 1; i++)
            {
                if (
                        (strTable[i, j] != "") && //пустая 
                        (strTable[i, j] != Const.NullTextReplace) && // для замены пустых 
                        (strTable[i, j] != Const.ExcelTextCellAsBblock) && //  [Блок]
                        (strTable[i, j] != Const.ExcelTextCellAsAttribut) // [Атрибут]
                    //excelCellBlockText.TextValue=
                    )
                {
                    listBlocks.Add(new ExcelRangeText
                    {
                        TextValue = strTable[i, j],
                        RowCell = i,
                        ColumnCell=j
                    }) ;
                }
            }

            return listBlocks;
            // throw new NotImplementedException();
        }

        // Список атр.
        public List<ExcelRangeText> GetExcelRangeAttribute()
        {
            // Ищем "[Атрибут]"
            excelCellAttributeText = SearchValueInArray.GetCellCoordinatesInArray(Const.ExcelTextCellAsAttribut, strTable);

            throw new NotImplementedException();
        }

        // Список данных для выгрузки. Рабочий метод!
        public List<BlockData> GetListBlockDataToPush()
        {
            throw new NotImplementedException();
        }


        #region ToString Override!
        /// <summary>
        /// Список данных для выгрузки. Тест!!
        /// </summary>
        /// <param name="ss">Липовый параметр, стобы перегрузить основной метод</param>
        /// <returns>Заранее заданный список, используется в ToString!</returns>
        public List<BlockData> GetListBlockDataToPush(string ss)
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

        /// <summary>
        /// переопределим ToString.
        /// </summary>
        /// <returns>результат, т.е. то , что будет передаваться в акад.</returns>
        public override string ToString()
        {
            string  str="";

            foreach (var BlData in GetListBlockDataToPush("ss"))
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
        #endregion

    }
}
