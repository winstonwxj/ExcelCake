using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;

namespace ExcelCake.Intrusive
{
    public class ExportExcelSetting
    {
        public List<ExportColumn> ExportColumns { set; get; }
        public ExportStyle ExportStyle { set; get; }

        public ExportExcelSetting(Type type)
        {
            ExportStyle = new ExportStyle();
            ExportColumns = new List<ExportColumn>();
            if (type == null)
            {
                return;
            }

            #region 组织表头
            //var modalType = EnumExportModalType.ALL;
            var classAttrArry = type.GetCustomAttributes(typeof(ExportModalAttribute), true);
            if (classAttrArry != null && classAttrArry.Length > 0)
            {
                //modalType = ((ExportModalAttribute)classAttrArry[0]).ExportModal;
                ExportStyle.Title = ((ExportModalAttribute)classAttrArry[0]).Title;
                //exportSetting.ExportStyle.HeadColor = ((ExportModalAttribute)classAttrArry[0]).HeadColor;
            }

            //导出字段
            var properties = type.GetProperties();

            foreach (var proper in properties)
            {
                var noexportAttrArry = proper.GetCustomAttributes(typeof(NoExportAttribute), true);
                if (noexportAttrArry == null || noexportAttrArry.Length == 0)
                {
                    var column = new ExportColumn(proper);
                    if (column != null)
                    {
                        ExportColumns.Add(column);
                    }
                }
            }
            #endregion

            #region 排序
            ExportColumns.Sort((a, b) => a.Index.CompareTo(b.Index));
            #endregion
        }
    }
}
