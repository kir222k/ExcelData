﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelData
{
    public class CreatePullData
    {
        public void DataToAcad()
        {
            EpPlusExcel ED = new EpPlusExcel();
            string[,] strTable = ED.GetDataExel();

            // найдем элемент "Блок"

            // найдем элемент "Атрибут:"

            // пойдем в столбце "Блок"а от "Блок"а вниз до 1го имени блока.
            // имена блоков - запомним 
            // в список БЛК (<имя блока1>.<x1,y1> <имя блока2>.<x2,y2> ...).

            // пройдем по строке "Атрибут:"ов и найдем столбцы, где есть значения -
            // имена атрибутов - запомним 
            // в список АТР (<имя атр1>.<x1,y1> <имя атр2>.<x2,y2> ...).

            // пойдем по списку БЛК - составим новый список:
            // БАТ (<имя блока1>.<список атрибутов1> <имя блока2>.<список атрибутов2>...)
            
            
            // БАТ список обрабатываем в др. молуле и заполняем вхеждения блоков в кад файле



        }

    }
}
