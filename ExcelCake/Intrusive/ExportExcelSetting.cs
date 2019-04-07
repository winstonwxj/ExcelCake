using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Drawing;

namespace ExcelCake.Intrusive
{
    public class ExportExcelSetting
    {
        public List<ExportColumn> ExportColumns { set; get; }
        public ExportStyle ExportStyle { set; get; }
        public List<KeyValuePair<string,int>> MergeList { set; get; }

        private ExportExcelSetting()
        {

        }

        public ExportExcelSetting(Type type)
        {
            ExportStyle = new ExportStyle();
            ExportColumns = new List<ExportColumn>();
            MergeList = new List<KeyValuePair<string, int>>();
            if (type == null)
            {
                return;
            }

            #region 组织表头
            var classAttrArry = type.GetCustomAttributes(typeof(ExportEntityAttribute), true);
            if (classAttrArry == null || classAttrArry.Length == 0)
            {
                return;
            }

            var exportEntity = (ExportEntityAttribute)classAttrArry[0];
            ExportStyle.Title = exportEntity.Title;
            ExportStyle.HeadColor = ColorTranslator.FromHtml(exportEntity.HeadColor.ToString());
            ExportStyle.TitleColor = ColorTranslator.FromHtml(exportEntity.TitleColor.ToString());
            ExportStyle.TitleFontSize = exportEntity.TitleFontSize;
            ExportStyle.IsTitleBold = exportEntity.IsTitleBold;
            ExportStyle.TitleColumnSpan = exportEntity.TitleColumnSpan;
            ExportStyle.HeadFontSize = exportEntity.HeadFontSize;
            ExportStyle.IsHeadBold = exportEntity.IsHeadBold;
            ExportStyle.ContentColor = ColorTranslator.FromHtml(exportEntity.ContentColor.ToString());
            ExportStyle.ContentFontSize = exportEntity.ContentFontSize;
            ExportStyle.IsContentBold = exportEntity.IsContentBold;
            var properties = type.GetProperties();

            foreach (var proper in properties)
            {
                var exportAttrArry = proper.GetCustomAttributes(typeof(ExportAttribute), true);
                if (exportAttrArry != null && exportAttrArry.Length > 0)
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
            //ExportColumns.Sort((a, b) => a.Index.CompareTo(b.Index));
            var groupTemp = ExportColumns.GroupBy(o => o.MergeText);
            ExportColumns = new List<ExportColumn>();
            foreach(var item in groupTemp)
            {
                var mergeList = item.ToList();
                if (mergeList.Count == 1)
                {
                    mergeList.First().MergeText = "";
                }
                else
                {
                    mergeList.Sort((a, b) => a.Index.CompareTo(b.Index));
                    MergeList.Add(new KeyValuePair<string, int>(item.Key, item.Count()));
                }
                
                ExportColumns.AddRange(mergeList);
            }
            #endregion
        }
    }
}
