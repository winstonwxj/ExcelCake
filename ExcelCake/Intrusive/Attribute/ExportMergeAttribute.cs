using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelCake.Intrusive
{
    /// <summary>
    /// 导出列合并特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class ExportMergeAttribute: Attribute
    {
        private string _MergeText { get; set; }

        /// <summary>
        /// 合并后文本
        /// </summary>
        public string MergeText
        {
            get
            {
                return _MergeText;
            }
            set
            {
                _MergeText = value;
            }
        }

        public ExportMergeAttribute(string mergeText)
        {
            _MergeText = mergeText ?? "";
        }
    }
}
