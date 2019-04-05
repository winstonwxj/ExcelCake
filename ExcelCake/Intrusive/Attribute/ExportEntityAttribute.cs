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
        private EnumColor _TitleColor;
        
        private string _Title;

        private int _TitleFontSize;

        private bool _IsTitleBold;

        private int _TitleColumnSpan; 

        private EnumColor _HeadColor;

        private int _HeadFontSize;

        private bool _IsHeadBold;

        private EnumColor _ContentColor;

        private int _ContentFontSize;

        private bool _IsContentBold;

        //列高

        //是否自动换行

        /// <summary>
        /// 标题背景颜色
        /// </summary>
        public EnumColor TitleColor
        {
            get
            {
                return _TitleColor;
            }
        }

        /// <summary>
        /// 标题文本
        /// </summary>
        public string Title
        {
            get
            {
                return _Title;
            }
        }

        /// <summary>
        /// 标题文本字号
        /// </summary>
        public int TitleFontSize
        {
            get
            {
                return _TitleFontSize;
            }
        }

        /// <summary>
        /// 标题是否粗体
        /// </summary>
        public bool IsTitleBold
        {
            get
            {
                return _IsTitleBold;
            }
        }

        /// <summary>
        /// 标题合并列数
        /// </summary>
        public int TitleColumnSpan
        {
            get
            {
                return _TitleColumnSpan;
            }
        }

        /// <summary>
        /// 列头背景颜色
        /// </summary>
        public EnumColor HeadColor
        {
            get
            {
                return _HeadColor;
            }
        }

        /// <summary>
        /// 列头字号
        /// </summary>
        public int HeadFontSize
        {
            get
            {
                return _HeadFontSize;
            }
        }

        /// <summary>
        /// 列头是否粗体
        /// </summary>
        public bool IsHeadBold
        {
            get
            {
                return _IsHeadBold;
            }
        }

        /// <summary>
        /// 内容背景色
        /// </summary>
        public EnumColor ContentColor
        {
            get
            {
                return _ContentColor;
            }
        }

        /// <summary>
        /// 内容字号
        /// </summary>
        public int ContentFontSize
        {
            get
            {
                return _ContentFontSize;
            }
        }

        /// <summary>
        /// 内容是否粗体
        /// </summary>
        public bool IsContentBold
        {
            get
            {
                return _IsContentBold;
            }
        }

        private ExportEntityAttribute()
        {

        }

        public ExportEntityAttribute(string title = "", int titleFontSize = 14, int headFontSize = 12, int contentFontSize = 10, bool isTitleBold = true, bool isHeadBold = true, bool isContentBold = false,int titleColumnSpan=0)
        {
            this._TitleColor = EnumColor.White;
            this._HeadColor = EnumColor.White;
            this._ContentColor = EnumColor.White;
            this._Title = title;
            this._TitleFontSize = titleFontSize;
            this._HeadFontSize = headFontSize;
            this._ContentFontSize = contentFontSize;
            this._IsTitleBold = isTitleBold;
            this._IsHeadBold = isHeadBold;
            this._IsContentBold = isContentBold;
            this._TitleColumnSpan = titleColumnSpan;
        }

        public ExportEntityAttribute(EnumColor headColor, string title="", int titleFontSize = 14, int headFontSize = 12, int contentFontSize = 10, bool isTitleBold = true, bool isHeadBold = true, bool isContentBold = false, int titleColumnSpan = 0)
        {
            this._TitleColor = EnumColor.White;
            this._HeadColor = headColor;
            this._ContentColor = EnumColor.White;
            this._Title = title;
            this._TitleFontSize = titleFontSize;
            this._HeadFontSize = headFontSize;
            this._ContentFontSize = contentFontSize;
            this._IsTitleBold = isTitleBold;
            this._IsHeadBold = isHeadBold;
            this._IsContentBold = isContentBold;
            this._TitleColumnSpan = titleColumnSpan;
        }

        public ExportEntityAttribute(EnumColor titleColor, EnumColor headColor, string title="", int titleFontSize = 14, int headFontSize = 12, int contentFontSize = 10, bool isTitleBold = true, bool isHeadBold = true, bool isContentBold = false, int titleColumnSpan = 0)
        {
            this._TitleColor = titleColor;
            this._HeadColor = headColor;
            this._ContentColor = EnumColor.White;
            this._Title = title;
            this._TitleFontSize = titleFontSize;
            this._HeadFontSize = headFontSize;
            this._ContentFontSize = contentFontSize;
            this._IsTitleBold = isTitleBold;
            this._IsHeadBold = isHeadBold;
            this._IsContentBold = isContentBold;
            this._TitleColumnSpan = titleColumnSpan;
        }

        public ExportEntityAttribute(EnumColor titleColor, EnumColor headColor, EnumColor contentColor, string title="", int titleFontSize = 14, int headFontSize = 12, int contentFontSize = 10, bool isTitleBold = true, bool isHeadBold = true, bool isContentBold = false, int titleColumnSpan = 0)
        {
            this._TitleColor = titleColor;
            this._HeadColor = headColor;
            this._ContentColor = contentColor;
            this._Title = title;
            this._TitleFontSize = titleFontSize;
            this._HeadFontSize = headFontSize;
            this._ContentFontSize = contentFontSize;
            this._IsTitleBold = isTitleBold;
            this._IsHeadBold = isHeadBold;
            this._IsContentBold = isContentBold;
            this._TitleColumnSpan = titleColumnSpan;
        }
    }
}