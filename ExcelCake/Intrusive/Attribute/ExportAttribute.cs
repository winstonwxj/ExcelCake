using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExcelCake.Intrusive
{
    /// <summary>
    /// 导出特性，标注导出的属性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class ExportAttribute:Attribute
    {
        /// <summary>
        /// 导出名称
        /// </summary>
        private string name;

        /// <summary>
        /// 导出排序索引
        /// </summary>
        private int sortIndex;

        public string Name
        {
            get
            {
                return name;
            }
        }

        public int SortIndex
        {
            get
            {
                return sortIndex;
            }
        }

        //合并相同列(list中)

        //prefix

        //suffix

        //WrapMode(自动列宽，自动换行)

        //列宽

        //数据校验

        private ExportAttribute()
        {

        }

        public ExportAttribute(string name,int index=0)
        {
            this.name = name;
            sortIndex = index;
        }
    }
}