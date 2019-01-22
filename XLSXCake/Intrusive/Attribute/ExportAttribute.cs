using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace XLSXCake.Intrusive
{
    /// <summary>
    /// 导出特性，标注导出的字段
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class ExportAttribute:Attribute
    {
        public ExportAttribute()
        {

        }
    }
}