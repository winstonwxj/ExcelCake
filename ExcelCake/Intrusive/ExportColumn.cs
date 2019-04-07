using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace ExcelCake.Intrusive
{
    public class ExportColumn
    {
        public string Text { set; get; }
        public string Value { set; get; }
        public int Index { set; get; }
        public string Prefix { set; get; }
        public string Suffix { set; get; }
        public string MergeText { set; get; }

        public ExportColumn()
        {

        }

        public ExportColumn(PropertyInfo property)
        {
            if (property == null)
            {
                return;
            }
            var exportAttrArry = property.GetCustomAttributes(typeof(ExportAttribute), true);
            if (exportAttrArry != null && exportAttrArry.Length > 0)
            {
                var export = ((ExportAttribute)exportAttrArry[0]);
                string displayName = property.Name;
                if (!string.IsNullOrEmpty(export.Name))
                {
                    displayName = export.Name;
                }
                Text = displayName;
                Value = property.Name;
                Index = export.SortIndex;
                Prefix = export.Prefix;
                Suffix = export.Suffix;
                var mergeAttrArry = property.GetCustomAttributes(typeof(ExportMergeAttribute), true);
                if (mergeAttrArry != null && mergeAttrArry.Length > 0)
                {
                    MergeText = ((ExportMergeAttribute)mergeAttrArry[0]).MergeText;
                }
            }
        }
    }
}
