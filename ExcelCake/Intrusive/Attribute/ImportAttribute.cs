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
        /// <summary>
        /// 导入名称
        /// </summary>
        private string _Name;

        /// <summary>
        /// 导出名称
        /// </summary>
        public string Name
        {
            get
            {
                return _Name;
            }
        }

        private ImportAttribute()
        {

        }

        public ImportAttribute(string name)
        {
            _Name = name;
        }
    }
}
