using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Web;

namespace ExcelCake.Intrusive
{
    /// <summary>
    // 导出表头合并特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class ExportMergeHeadAttribute : Attribute
    {
        /// <summary>
        /// 合并文本
        /// </summary>
        private string text;

        public string Text
        {
            get
            {
                return text;
            }
        }

        public ExportMergeHeadAttribute(string mergetext)
        {
            text = mergetext;
        }
    }
}