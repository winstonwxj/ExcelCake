using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelCake.Intrusive;

namespace ExcelCake.Example
{
    [ExportEntity("用户信息")]
    public class UserInfo: ExcelBase
    {
        [Export("编号",1)]
        public int ID { set; get; }

        [Export("姓名", 2)]
        public string Name { set; get; }

        [Export("性别", 3)]
        public string Sex { set; get; }

        [Export("年龄", 4)]
        public int? Age { set; get; }

        [Export("电子邮件", 5)]
        public string Email { set; get; }

        [NoExport]
        public string TelPhone { set; get; }
    }
}
