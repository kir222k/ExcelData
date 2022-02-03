using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelData.Model
{
    public struct AttrData
    {
        public string AttributeTag;
        public string AttributeValue;
    }

    public struct BlockData
    {
        public string BlockName;
        public List<AttrData> ListAttributes;
    }

    public struct ExcelRangeText
    {
        public string TextValue;
        public int RowCell;
        public int ColumnCell;
    }
}
