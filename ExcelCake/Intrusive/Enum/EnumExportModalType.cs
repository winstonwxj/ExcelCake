using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExcelCake.Intrusive
{
    /// <summary>
    /// 导出模式类型枚举类
    /// </summary>
    public enum EnumExportModalType
    {
        /// <summary>
        /// 导出所有字段
        /// </summary>
        ALL = 0,
        /// <summary>
        /// 部分字段（导出ExportAttribute标注的字段）
        /// </summary>
        PART = 1,
        /// <summary>
        /// 反选(不导出NoExportAttribute标注的字段)
        /// </summary>
        INVERT = 2
    }
}