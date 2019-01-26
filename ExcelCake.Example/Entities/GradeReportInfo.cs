using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelCake.Example
{
    public class GradeReportInfo
    {
        public string ReportTitle { set; get; }
        /// <summary>
        /// 各班级各科目及格人数列表
        /// </summary>
        public List<ClassInfo> List1 { set; get; }
        /// <summary>
        /// 各班级各科目平均成绩
        /// </summary>
        public List<ClassInfo> List2 { set; get; }
        /// <summary>
        /// 各班级总分情况
        /// </summary>
        public List<ClassInfo> List3 { set; get; }

        public GradeReportInfo()
        {
            List1 = new List<ClassInfo>();
            List2 = new List<ClassInfo>();
            List3 = new List<ClassInfo>();
        }
    }
}
