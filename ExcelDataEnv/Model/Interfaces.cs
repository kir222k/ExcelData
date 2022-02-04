using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelData.Model
{
    interface IPullPushData

    {
        /// <summary>
        /// найдем тескт   "Блок" на листе Excel.
        /// </summary>
        /// <returns>значение типа ExcelRangeText</returns>
       // ExcelRangeText GetExcelCellBlockText();

        /// <summary>
        /// найдем тескт   "Блок" на листе Excel.
        /// </summary>
        /// <returns>значение типа ExcelRangeText</returns>
       // ExcelRangeText GetExcelCellAttributeText();

        /// <summary>
        /// пойдем в столбце "Блок"а от "Блок"а вниз до 1го имени блока.
        /// в список БЛК (имя блока1.x1,y1 .. имя блока2.x2,y ...).
        /// </summary>
        /// <returns>список типа ExcelRangeText</returns>
        List<ExcelRangeText> GetExcelRangeBlock();

        /// <summary>
        /// пройдем по строке "Атрибут:"ов и найдем столбцы, где есть значения -
        /// имена атрибутов - запомним 
        /// в список АТР (имя атр1.x1,y1 .. имя атр2.x2,y2 ...).
        /// </summary>
        /// <returns>список типа ExcelRangeText</returns>
        List<ExcelRangeText> GetExcelRangeAttribute();

        /// <summary>
        /// пойдем по списку БЛК - составим новый список:
        /// БАТ (имя блока1.список атрибутов1 имя блока2.список атрибутов2...)
        ///   берем из БАТ имя  блока и строку, где он сидит (х)
        ///   берем из АТР имя атрибута, его столбец (у)
        ///   записываем в БАТ:
        ///   1 элемент в списке:
        ///   имя блока, в его строке (x - а мы его знаем из БЛК) ищем столбец (у) атрибута -
        ///   получаем зн. атрибута.
        ///   добавляем к имени блока пару Имя атрибута.значение. 1й атрибут получили!
        ///   далее - получем 2й, 3й и .тд., пока не закончится список АТР.
        ///   посторяем для след. блока в списке БЛК
        /// </summary>
        /// <returns></returns>
        List<BlockData> GetListBlockDataToPush();


    }
}
