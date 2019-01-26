using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelCake.Example
{
    public class ClassInfo
    {
        public string ClassName { set; get; }
        public int PassCountSubject1 { set; get; }
        public int PassCountSubject2 { set; get; }
        public int PassCountSubject3 { set; get; }
        public int PassCountSubject4 { set; get; }
        public int PassCountSubject5 { set; get; }

        public double ScoreAvgSubject1 { set; get; }
        public double ScoreAvgSubject2 { set; get; }
        public double ScoreAvgSubject3 { set; get; }
        public double ScoreAvgSubject4 { set; get; }
        public double ScoreAvgSubject5 { set; get; }

        public double ScoreTotalMax { set; get; }
        public double ScoreTotalAvg { set; get; }
        public double ScoreTotalPassRate { set; get; }
    }
}
