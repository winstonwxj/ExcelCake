using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Web;

namespace ExcelCake.Intrusive
{
    /// <summary>
    /// 导出模式特性，标注类的导出模式
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    public class ExportModalAttribute:Attribute
    {
        /// <summary>
        /// 导出模式，默认为全部
        /// </summary>
        //private EnumExportModalType exportModal;
        /// <summary>
        /// 表头颜色
        /// </summary>
        private Color headColor;
        /// <summary>
        /// 表格标题
        /// </summary>
        private string title;

        //指定排序属性

        //列高

        //是否自动换行

        //public EnumExportModalType ExportModal
        //{
        //    get
        //    {
        //        return exportModal;
        //    }
        //}

        public Color HeadColor
        {
            get
            {
                return headColor;
            }
        }

        public string Title
        {
            get
            {
                return title;
            }
        }

        public ExportModalAttribute()
        {
            //exportModal = EnumExportModalType.ALL;
            headColor = Color.FromArgb(192, 192, 192);
            title = "";
        }

        public ExportModalAttribute(string title="")
        {
            //exportModal = modalType;
            headColor = Color.FromArgb(192, 192, 192);
            this.title = title;
        }

        public ExportModalAttribute(Color headColor,string title="")
        {
            //exportModal = modalType;
            this.headColor = headColor;
            this.title = title;
        }
    }
}