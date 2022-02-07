using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelData.Model;

namespace ExcelData.Class
{
    static class SearchValueInArray
    {
        /// <summary>
        /// Возращает Текст, i, j искомой ячеки
        /// </summary>
        /// <param name="cellValue"> Текст искомой ячейки, д. быть уникальным на листе</param>
        /// <param name="strTable"> 2 мерный массив строк - представление листа</param>
        /// <returns></returns>
        public static  ExcelRangeText GetCellCoordinatesInArray(string cellValue, string[,] strTable)
        {
            ExcelRangeText eT = new ExcelRangeText();
            //eT = null;

            // начнем искать.
            // берем массив, идем 
            int rows = strTable.GetUpperBound(0) + 1;    // количество строк
            int columns = strTable.GetUpperBound(1) + 1; // количество столбцов

            for (int i = 0; i < rows - 1; i++)
            {
                for (int j = 0; j < columns - 1; j++)
                {
                    //Console.Write(strTable[i, j].ToString() + " ");
                    if (strTable[i, j] == cellValue) 
                    {
                        eT.TextValue = cellValue;
                        eT.RowCell = i;
                        eT.ColumnCell = j;
                    }
                }
                //Console.WriteLine("\n");
            }
            return eT;
        }
    }
}
