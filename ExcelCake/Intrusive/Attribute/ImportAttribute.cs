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
        private bool _IsUseTempField;
        private string _TempField;
        private string _Prefix;
        private string _Suffix;
        private string _DataVerReg;
        private bool _IsRegFailThrowException;

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
        /// 是否使用临时字段
        /// </summary>
        public bool IsUseTempField
        {
            get
            {
                return _IsUseTempField;
            }
            set
            {
                _IsUseTempField = value;
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

        /// <summary>
        /// 前缀
        /// </summary>
        public string Prefix
        {
            get
            {
                return _Prefix;
            }
            set
            {
                _Prefix = value;
            }
        }

        /// <summary>
        /// 后缀
        /// </summary>
        public string Suffix
        {
            get
            {
                return _Suffix;
            }
            set
            {
                _Suffix = value;
            }
        }

        /// <summary>
        /// 数据校验正则
        /// </summary>
        public string DataVerReg
        {
            get
            {
                return _DataVerReg;
            }
            set
            {
                _DataVerReg = value;
            }
        }

        /// <summary>
        /// 正则验证失败是否抛出异常
        /// </summary>
        public bool IsRegFailThrowException
        {
            get
            {
                return _IsRegFailThrowException;
            }
            set
            {
                _IsRegFailThrowException = value;
            }
        }

        //数据校验

        private ImportAttribute()
        {
            _IsUseTempField = false;
        }

        public ImportAttribute(string name, string dataVerReg = "", bool isRegFailThrowException = false,string prefix="",string suffix="")
        {
            _Name = name??"";
            _IsUseTempField = false;
            _DataVerReg = dataVerReg??"";
            _IsRegFailThrowException = isRegFailThrowException;
            _Prefix = prefix??"";
            _Suffix = suffix ?? "";
        }

        public ImportAttribute(string name,bool isUseTempField, string tempField, string prefix = "", string suffix = "")
        {
            _Name = name??"";
            _IsUseTempField = isUseTempField;
            _TempField = tempField??"";
            _Prefix = prefix ?? "";
            _Suffix = suffix ?? "";
        }
    }
}
