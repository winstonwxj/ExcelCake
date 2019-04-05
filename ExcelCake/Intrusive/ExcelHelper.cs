using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelCake.Intrusive
{
    public abstract class ExcelBase
    {
        public string ExcelName { set; get; }
        public string SheetName { set; get; }
        public int RowIndex { set; get; }
        public int ColumnIndex { set; get; }
    }
}
