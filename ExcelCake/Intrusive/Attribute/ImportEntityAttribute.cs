using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelCake.Intrusive
{
    /// <summary>
    /// 导入特性，标注类的导入信息
    /// </summary>
    [AttributeUsage(AttributeTargets.Class)]
    public class ImportEntityAttribute: Attribute
    {
        private int _TitleRowIndex;
        private int _HeadRowIndex;
        private int _DataRowIndex;

        /// <summary>
        /// 标题行Index
        /// </summary>
        public int TitleRowIndex
        {
            get
            {
                return _TitleRowIndex;
            }
            set
            {
                _TitleRowIndex = value;
            }
        }

        /// <summary>
        /// 列头行Index
        /// </summary>
        public int HeadRowIndex
        {
            get
            {
                return _HeadRowIndex;
            }
            set
            {
                _HeadRowIndex = value;
            }
        }

        /// <summary>
        /// 数据行Index
        /// </summary>
        public int DataRowIndex
        {
            get
            {
                return _DataRowIndex;
            }
            set
            {
                _DataRowIndex = value;
            }
        }

        public ImportEntityAttribute()
        {
            _TitleRowIndex = 0;
            _HeadRowIndex = 1;
            _DataRowIndex = 2;
        }

        public ImportEntityAttribute(int headRowIndex, int dataRowIndex)
        {
            _TitleRowIndex = 0;
            _HeadRowIndex = headRowIndex < 1 ? 1 : headRowIndex;
            _DataRowIndex = dataRowIndex < 2 ? 2 : dataRowIndex;
        }

        public ImportEntityAttribute(int titleRowIndex,int headRowIndex,int dataRowIndex)
        {
            _TitleRowIndex = titleRowIndex < 0 ? 0 : titleRowIndex;
            _HeadRowIndex = headRowIndex < 1 ? 1 : headRowIndex;
            _DataRowIndex = dataRowIndex < 2 ? 2 : dataRowIndex;
        }
    }
}
