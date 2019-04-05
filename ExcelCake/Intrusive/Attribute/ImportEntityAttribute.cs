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
        private int _DataStartColumn;

        /// <summary>
        /// 数据起始行
        /// </summary>
        public int DataStartColumn
        {
            get
            {
                return _DataStartColumn;
            }
            set
            {
                _DataStartColumn = value;
            }
        }

        public ImportEntityAttribute()
        {
            _DataStartColumn = 1;
        }

        public ImportEntityAttribute(int dataColumnStart)
        {
            _DataStartColumn = dataColumnStart<1?1:dataColumnStart;
        }
    }
}
