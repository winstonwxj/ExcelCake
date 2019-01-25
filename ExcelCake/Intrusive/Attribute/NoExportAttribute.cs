using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace XLSXCake.Intrusive
{
    /// <summary>
    /// 不导出特性，标注不导出的字段
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class NoExportAttribute:Attribute
    {
        public NoExportAttribute()
        {

        }
    }
}