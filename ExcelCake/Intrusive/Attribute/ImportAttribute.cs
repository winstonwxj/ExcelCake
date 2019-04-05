using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelCake.Intrusive
{
    /// <summary>
    /// 导入特性，标注导入的属性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property)]
    public class ImportAttribute: Attribute
    {
        private string _Name;
        private bool _IsConvert;
        private string _TempField;

        /// <summary>
        /// 导入名称
        /// </summary>
        public string Name
        {
            get
            {
                return _Name;
            }
        }

        /// <summary>
        /// 是否需要转换
        /// </summary>
        public bool IsConvert
        {
            get
            {
                return _IsConvert;
            }
            set
            {
                _IsConvert = value;
            }
        }

        /// <summary>
        /// 临时字段
        /// </summary>
        public string TempField
        {
            get
            {
                return _TempField;
            }
            set
            {
                _TempField = value;
            }
        }

        private ImportAttribute()
        {
            _IsConvert = false;
        }

        public ImportAttribute(string name)
        {
            _Name = name;
            _IsConvert = false;
        }

        public ImportAttribute(string name,bool isConvert,string tempField)
        {
            _Name = name;
            _IsConvert = isConvert;
            _TempField = tempField;
        }
    }
}
