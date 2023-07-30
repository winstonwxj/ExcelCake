using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelCake.NoIntrusive
{
    [Serializable]
    internal abstract class TemplateSettingRange: TemplateSettingBase
    {
        internal List<TemplateSettingField> Fields { get; set; }

        internal string DataSource { set; get; }
        //public string Address { set; get; }
        internal string AddressLeftTop { set; get; }
        internal string AddressRightBottom { set; get; }

        internal int FromRow { set; get; }
        internal int FromCol { set; get; }
        internal int ToRow { set; get; }
        internal int ToCol { set; get; }

        internal abstract void Draw(ExcelWorksheet workSheet, ExcelObject dataSource);
    }
}
