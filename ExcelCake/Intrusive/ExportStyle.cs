using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace ExcelCake.Intrusive
{
    public class ExportStyle
    {
        /// <summary>
        /// 标题背景颜色
        /// </summary>
        public Color TitleColor { set; get; }

        /// <summary>
        /// 标题文本
        /// </summary>
        public string Title { set; get; }

        /// <summary>
        /// 标题文本字号
        /// </summary>
        public int TitleFontSize { set; get; }

        /// <summary>
        /// 标题是否粗体
        /// </summary>
        public bool IsTitleBold { set; get; }

        /// <summary>
        /// 标题合并列数
        /// </summary>
        public int TitleColumnSpan { set; get; }

        /// <summary>
        /// 列头背景颜色
        /// </summary>
        public Color HeadColor { set; get; }

        /// <summary>
        /// 列头字号
        /// </summary>
        public int HeadFontSize { set; get; }

        /// <summary>
        /// 列头是否粗体
        /// </summary>
        public bool IsHeadBold { set; get; }

        /// <summary>
        /// 内容背景色
        /// </summary>
        public Color ContentColor { set; get; }

        /// <summary>
        /// 内容字号
        /// </summary>
        public int ContentFontSize { set; get; }

        /// <summary>
        /// 内容是否粗体
        /// </summary>
        public bool IsContentBold { set; get; }
    }
}
