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
        public bool IsConvert { set; get; }
        public string TempField { set; get; }
        

        public ImportColumn()
        {
            IsConvert = false;
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
                IsConvert = import.IsConvert;
                TempField = import.TempField;
            }
        }
    }
}
