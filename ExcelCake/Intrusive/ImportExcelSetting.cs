using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelCake.Intrusive
{
    public class ImportExcelSetting
    {
        public List<ImportColumn> ImportColumns { set; get; }
        public ImportStyle ImportStyle { set; get; }

        private ImportExcelSetting()
        {

        }

        public ImportExcelSetting(Type type)
        {
            ImportStyle = new ImportStyle();
            ImportColumns = new List<ImportColumn>();
            if (type == null)
            {
                return;
            }

            #region 组织表头
            var classAttrArry = type.GetCustomAttributes(typeof(ImportEntityAttribute), true);
            if (classAttrArry == null || classAttrArry.Length == 0)
            {
                return;
            }

            var importEntity = (ImportEntityAttribute)classAttrArry[0];

            ImportStyle.TitleRowIndex = importEntity.TitleRowIndex;
            ImportStyle.HeadRowIndex = importEntity.HeadRowIndex;
            ImportStyle.DataRowIndex = importEntity.DataRowIndex;

            var properties = type.GetProperties();

            foreach (var proper in properties)
            {
                var importAttrArry = proper.GetCustomAttributes(typeof(ImportAttribute), true);
                if (importAttrArry != null && importAttrArry.Length > 0)
                {
                    var column = new ImportColumn(proper);
                    if (column != null)
                    {
                        ImportColumns.Add(column);
                    }
                }
            }
            #endregion
        }
    }
}
