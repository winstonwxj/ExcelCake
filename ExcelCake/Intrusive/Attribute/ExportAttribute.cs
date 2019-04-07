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
        private string _Name;
        private int _SortIndex;
        private string _Prefix;
        private string _Suffix;

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

        /// <summary>
        /// 前缀
        /// </summary>
        public string Prefix
        {
            get
            {
                return _Prefix;
            }
            set
            {
                _Prefix = value;
            }
        }

        /// <summary>
        /// 后缀
        /// </summary>
        public string Suffix
        {
            get
            {
                return _Suffix;
            }
            set
            {
                _Suffix = value;
            }
        }

        //合并相同列(list中)

        //WrapMode(自动列宽，自动换行)

        //列宽

        private ExportAttribute(string prefix="",string suffix="")
        {
            _Prefix = prefix ?? "";
            _Suffix = suffix ?? "";
        }

        public ExportAttribute(string name,int index=0, string prefix = "", string suffix = "")
        {
            _Name = name ?? "" ;
            _SortIndex = index;
            _Prefix = prefix ?? "";
            _Suffix = suffix ?? "";
        }
    }
}