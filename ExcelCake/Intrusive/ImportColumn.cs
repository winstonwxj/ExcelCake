using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace ExcelCake.Intrusive
{
    public class ImportColumn
    {
        public string Text { set; get; }
        public string Name { set; get; }
        public int ColumnIndex { set; get; }
        public bool IsUseTempField { set; get; }
        public string TempField { set; get; }
        public string Prefix { set; get; }
        public string Suffix { set; get; }
        public string DataVerReg { set; get; }
        public bool IsRegFailThrowException { set; get; }


        public ImportColumn()
        {
            IsUseTempField = false;
        }

        public ImportColumn(PropertyInfo property)
        {
            if (property == null)
            {
                return;
            }
            var importAttrArry = property.GetCustomAttributes(typeof(ImportAttribute), true);
            if (importAttrArry != null && importAttrArry.Length > 0)
            {
                var import = ((ImportAttribute)importAttrArry[0]);
                Name = property.Name;
                Text = import.Name;
                IsUseTempField = import.IsUseTempField;
                TempField = import.TempField;
                Prefix = import.Prefix;
                Suffix = import.Suffix;
                DataVerReg = import.DataVerReg;
                IsRegFailThrowException = import.IsRegFailThrowException;
            }
        }
    }
}
