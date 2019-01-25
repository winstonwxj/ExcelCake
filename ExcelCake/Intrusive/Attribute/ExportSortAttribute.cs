using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace XLSXCake.Intrusive
{
    /// <summary>
    /// 导出排序特性，标注导出字段顺序
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class ExportSortAttribute : Attribute
    {
        /// <summary>
        /// 导出排序索引
        /// </summary>
        private int sortIndex;

        public int SortIndex
        {
            get
            {
                return sortIndex;
            }
        }

        public ExportSortAttribute(int index)
        {
            sortIndex = index;
        }
    }
}