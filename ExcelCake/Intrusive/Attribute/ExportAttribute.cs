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
        private string _Name;

        /// <summary>
        /// 导出排序索引
        /// </summary>
        private int _SortIndex;

        /// <summary>
        /// 导出名称
        /// </summary>
        public string Name
        {
            get
            {
                return _Name;
            }
        }

        /// <summary>
        /// 导出排序索引
        /// </summary>
        public int SortIndex
        {
            get
            {
                return _SortIndex;
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
            _Name = name;
            _SortIndex = index;
        }
    }
}