using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelData
{
    public struct AttrData
    {
        public string AttributeTag;
        public string AttributeValue;
    }

    public class Group

    {
        // Имя блока
        AttrData Name { get; set; }

        // УЧАСТОК (секция ВРУ номер) - 1
        AttrData PartPanel { get; set; }

        // N.АПП1 (номер аппарата, QF) - QF1
        AttrData UnitNumber { get; set; }

        // Р.УСТ,КВТ (Pу) - 2,00
        AttrData PowerInstalled { get; set; }

        //

    }
}
