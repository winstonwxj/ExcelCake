using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Web;

namespace ExcelCake.Intrusive
{
    /// <summary>
    /// 导出特性，标注类的导出信息
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    public class ExportEntityAttribute:Attribute
    {
        private Color titleColor;
        
        private string title;

        private int titleFontSize;

        private bool isTitleBold;

        private int titleColumnSpan; 

        private Color headColor;

        private int headFontSize;

        private bool isHeadBold;

        private Color contentColor;

        private int contentFontSize;

        private bool isContentBold;

        //列高

        //是否自动换行

        /// <summary>
        /// 标题背景颜色
        /// </summary>
        public Color TitleColor
        {
            get
            {
                return titleColor;
            }
        }

        /// <summary>
        /// 标题文本
        /// </summary>
        public string Title
        {
            get
            {
                return title;
            }
        }

        /// <summary>
        /// 标题文本字号
        /// </summary>
        public int TitleFontSize
        {
            get
            {
                return titleFontSize;
            }
        }

        /// <summary>
        /// 标题是否粗体
        /// </summary>
        public bool IsTitleBold
        {
            get
            {
                return isTitleBold;
            }
        }

        /// <summary>
        /// 标题合并列数
        /// </summary>
        public int TitleColumnSpan
        {
            get
            {
                return titleColumnSpan;
            }
        }

        /// <summary>
        /// 列头背景颜色
        /// </summary>
        public Color HeadColor
        {
            get
            {
                return headColor;
            }
        }

        /// <summary>
        /// 列头字号
        /// </summary>
        public int HeadFontSize
        {
            get
            {
                return headFontSize;
            }
        }

        /// <summary>
        /// 列头是否粗体
        /// </summary>
        public bool IsHeadBold
        {
            get
            {
                return isHeadBold;
            }
        }

        /// <summary>
        /// 内容背景色
        /// </summary>
        public Color ContentColor
        {
            get
            {
                return contentColor;
            }
        }

        /// <summary>
        /// 内容字号
        /// </summary>
        public int ContentFontSize
        {
            get
            {
                return contentFontSize;
            }
        }

        /// <summary>
        /// 内容是否粗体
        /// </summary>
        public bool IsContentBold
        {
            get
            {
                return isContentBold;
            }
        }

        private ExportEntityAttribute()
        {

        }

        public ExportEntityAttribute(string title = "", int titleFontSize = 14, int headFontSize = 12, int contentFontSize = 10, bool isTitleBold = true, bool isHeadBold = true, bool isContentBold = false,int titleColumnSpan=1)
        {
            this.titleColor = Color.White;
            this.headColor = Color.White;
            this.contentColor = Color.White;
            this.title = title;
            this.titleFontSize = titleFontSize;
            this.headFontSize = headFontSize;
            this.contentFontSize = contentFontSize;
            this.isTitleBold = isTitleBold;
            this.isHeadBold = isHeadBold;
            this.isContentBold = isContentBold;
            this.titleColumnSpan = titleColumnSpan;
        }

        public ExportEntityAttribute(Color headColor, string title="", int titleFontSize = 14, int headFontSize = 12, int contentFontSize = 10, bool isTitleBold = true, bool isHeadBold = true, bool isContentBold = false, int titleColumnSpan = 1)
        {
            this.titleColor = Color.White;
            this.headColor = headColor;
            this.contentColor = Color.White;
            this.title = title;
            this.titleFontSize = titleFontSize;
            this.headFontSize = headFontSize;
            this.contentFontSize = contentFontSize;
            this.isTitleBold = isTitleBold;
            this.isHeadBold = isHeadBold;
            this.isContentBold = isContentBold;
            this.titleColumnSpan = titleColumnSpan;
        }

        public ExportEntityAttribute(Color titleColor,Color headColor, string title="", int titleFontSize = 14, int headFontSize = 12, int contentFontSize = 10, bool isTitleBold = true, bool isHeadBold = true, bool isContentBold = false, int titleColumnSpan = 1)
        {
            this.titleColor = titleColor;
            this.headColor = headColor;
            this.contentColor = Color.White;
            this.title = title;
            this.titleFontSize = titleFontSize;
            this.headFontSize = headFontSize;
            this.contentFontSize = contentFontSize;
            this.isTitleBold = isTitleBold;
            this.isHeadBold = isHeadBold;
            this.isContentBold = isContentBold;
            this.titleColumnSpan = titleColumnSpan;
        }

        public ExportEntityAttribute(Color titleColor, Color headColor,Color contentColor, string title="", int titleFontSize = 14, int headFontSize = 12, int contentFontSize = 10, bool isTitleBold = true, bool isHeadBold = true, bool isContentBold = false, int titleColumnSpan = 1)
        {
            this.titleColor = titleColor;
            this.headColor = headColor;
            this.contentColor = contentColor;
            this.title = title;
            this.titleFontSize = titleFontSize;
            this.headFontSize = headFontSize;
            this.contentFontSize = contentFontSize;
            this.isTitleBold = isTitleBold;
            this.isHeadBold = isHeadBold;
            this.isContentBold = isContentBold;
            this.titleColumnSpan = titleColumnSpan;
        }
    }
}