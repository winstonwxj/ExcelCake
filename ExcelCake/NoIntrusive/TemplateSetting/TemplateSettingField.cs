using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;

namespace ExcelCake.NoIntrusive
{
    [Serializable]
    internal class TemplateSettingField : TemplateSettingBase
    {
        public string Field { set; get; }

        private TemplateSettingField()
        {

        }

        internal static TemplateSettingField Create(ExcelRangeBase cell)
        {
            var entity = new TemplateSettingField();
            entity.CurrentCell = cell;
            entity.Content = cell.Value?.ToString() ?? "";
            entity.AnalyseSetting();
            return entity;
        }

        protected override void AnalyseSetting()
        {
            var cellValue = this.Content;

            var arry = cellValue.Split(new char[2] { '{', '}' }, StringSplitOptions.RemoveEmptyEntries);
            if (arry.Length == 0)
            {
                return;
            }
            foreach (var item in arry)
            {
                if (item.IndexOf(":") > -1 && item.IndexOf(";") > -1)
                {
                    var settingItemArry = item.Split(new char[1] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                    if (settingItemArry.Length == 0)
                    {
                        continue;
                    }

                    foreach (var arryItem in settingItemArry)
                    {
                        var settingItem = arryItem.Split(':');
                        if (settingItem.Length < 2)
                        {
                            continue;
                        }
                        var key = settingItem[0];
                        var value = settingItem[1];
                        if (string.IsNullOrEmpty(key))
                        {
                            continue;
                        }

                        switch (key.ToUpper())
                        {
                            case "FIELD":
                                {
                                    this.Field = value.ToUpper();
                                }
                                break;
                        }
                    }

                }
            }

            this.ScopeRange = this.CurrentCell;
        }
    }
}
