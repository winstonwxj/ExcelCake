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
        public Color TitleColor;

        /// <summary>
        /// 标题文本
        /// </summary>
        public string Title { set; get; }

        /// <summary>
        /// 标题文本字号
        /// </summary>
        public int TitleFontSize;

        /// <summary>
        /// 标题是否粗体
        /// </summary>
        public bool IsTitleBold;

        /// <summary>
        /// 标题合并列数
        /// </summary>
        public int TitleColumnSpan;

        /// <summary>
        /// 列头背景颜色
        /// </summary>
        public Color HeadColor { set; get; }

        /// <summary>
        /// 列头字号
        /// </summary>
        public int HeadFontSize;

        /// <summary>
        /// 列头是否粗体
        /// </summary>
        public bool IsHeadBold;

        /// <summary>
        /// 内容背景色
        /// </summary>
        public Color ContentColor;

        /// <summary>
        /// 内容字号
        /// </summary>
        public int ContentFontSize;

        /// <summary>
        /// 内容是否粗体
        /// </summary>
        public bool IsContentBold;
    }
}
