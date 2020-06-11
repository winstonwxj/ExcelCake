using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelCake.NoIntrusive
{
    [Serializable]
    internal abstract class TemplateSettingBase
    {
        internal string Content { set; get; }

        /// <summary>
        /// 配置所在单元格
        /// </summary>
        internal ExcelRangeBase CurrentCell { set; get; }

        /// <summary>
        /// 作用范围
        /// </summary>
        internal ExcelRangeBase ScopeRange { set; get; }

        protected abstract void AnalyseSetting();
    }
}
