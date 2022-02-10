/* Кирилл Уваров 2022г. 10 февраля. u.k.send@gmail.com. +79062644029
*/

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
            // берем массив и ищем в столбце excelCellBlockText.ColumnCell
            int j = excelCellBlockText.ColumnCell;
            for (int i = 0; i < rows - 1; i++)
            {
                if (IsValueCorrect(strTable[i, j]))
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
            var listAttrs = new List<ExcelRangeText>();
            // Ищем "[Атрибут]"
            excelCellAttributeText = SearchValueInArray.GetCellCoordinatesInArray(Const.ExcelTextCellAsAttribute, strTable);

            int columns = strTable.GetUpperBound(1) + 1; // количество столбцов
            // берем массив и ищем в строке excelCellAttributeText.RowCell
            int i = excelCellAttributeText.RowCell;
            for (int j = 0; j < columns - 1; j++)
            {
                if (IsValueCorrect(strTable[i, j]))
                {
                    listAttrs.Add(new ExcelRangeText
                    {
                        TextValue = strTable[i, j],
                        RowCell = i,
                        ColumnCell = j
                    });
                }
            }

            return listAttrs;
            //throw new NotImplementedException();
        }

        // Список данных для выгрузки. Оновной метод!
        public List<BlockData> GetListBlockDataToPush()
        {


            var listBlockData = new List<BlockData>();
            // спимок блоков
            List<ExcelRangeText> listBlocks = new();
            listBlocks = GetExcelRangeBlock();

            // список атрибутов
            List<ExcelRangeText> listAttrs = new ();
            listAttrs = GetExcelRangeAttribute();

            // итак, по списку блоков
            foreach (ExcelRangeText bl in listBlocks)
            {
                // создадим список пар "АтрибутТэг.АтрибутЗнач"
                List<AttrData> listAttData = new List<AttrData>();

                // пройдем по списку атрибутов
                foreach (ExcelRangeText att in listAttrs)
                {
                    // заберем тэг атрибута
                    var attDateElement = new AttrData();
                    attDateElement.AttributeTag = att.TextValue;
                    // заберем его значение - на пересеч. строки блока и столбца атрибута
                    // строка блока
                    int blockRow = bl.RowCell;
                    // столбец атрибута
                    int attColumn = att.ColumnCell;
                    // значение атрибута с таким тэгом для данного вхождения блока
                    string attValue = strTable[blockRow, attColumn];
                    // проверим, не пустое ли значение атрибута
                    if (attValue != Const.NullTextReplace) // если не равно <emp>
                    {
                        attDateElement.AttributeValue = attValue;
                    }
                    else // Если равно <emp>
                    {
                        attDateElement.AttributeValue = string.Empty; //  заменим на объект Empty? да!
                    }

                    // добавим "АтрибутТэг.АтрибутЗнач" в список
                    listAttData.Add(attDateElement);
                }

                //listAttData = new List<AttrData>
                //{
                //    new AttrData
                //    {
                //        AttributeTag="Tag",
                //        AttributeValue="QF1"
                //    }
                //};

                // создадим элемент типа BlockData
                BlockData blData = new BlockData
                {
                    // добавим имя блока
                    BlockName=bl.TextValue,
                    // добавим список пар "АтрибутТэг.АтрибутЗнач"
                    ListAttributes= listAttData
                };

                // добавим к списку элемент типа BlockData:
                listBlockData.Add(blData);
            }



            return listBlockData;
            //throw new NotImplementedException();
        }


        private bool IsValueCorrect (string str)
        {
            bool isCorrect = false;

            if (
                       (str != "") && //пустая 
                       (str != Const.NullTextReplace) && // для замены пустых 
                       (str != Const.ExcelTextCellAsBblock) && //  [Блок]
                       (str != Const.ExcelTextCellAsAttribute) // [Атрибут]
                                                                          //excelCellBlockText.TextValue=
                   )
            {
                isCorrect = true;
            }


                return isCorrect;
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

            //foreach (var BlData in GetListBlockDataToPush("ss"))
            foreach (var BlData in GetListBlockDataToPush())
            {
                str += "\n.\n.\nБлок: " + BlData.BlockName;
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
        /*
            Person person = new Person { Name = "John", Age = 12 };
            Console.WriteLine(person);
            // Output:
            // Person: John 12
        */
        #endregion

    }
}
