using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelData.Model
{

    public class BlockDataContainer

    {
        // Имя блока
        public string Name { get; set; }

        // УЧАСТОК (секция ВРУ номер) - 1
        public AttrData PartPanel { get; set; }

        // N.АПП1 (номер аппарата, QF) - QF1
        public AttrData UnitNumber { get; set; }

        // Р.УСТ,КВТ (Pу) - 2,00
        public AttrData PowerInstalled { get; set; }

        public override string ToString()
        {
            return "Данные блока: " + 
                "\nИмя: " + Name +
                "\nУчасток: " + PartPanel+
                "\nНомер аппарата: " + UnitNumber +
                "\nУстановленная мощность: "+ PowerInstalled +"кВт";

        }

    }
}
