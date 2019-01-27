using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelCake.NoIntrusive
{
    internal class TemplateSettingRange
    {
        public string Type { set; get; }
        public string DataSource { set; get; }
        //public string Address { set; get; }
        public string AddressLeftTop { set; get; }
        public string AddressRightBottom { set; get; }
        public string Field { set; get; }
        public string SettingString { set; get; }
        public ExcelRangeBase CurrentCell { set; get; }
        public int FromRow { set; get; }
        public int FromCol { set; get; }
        public int ToRow { set; get; }
        public int ToCol { set; get; }
    }
}
