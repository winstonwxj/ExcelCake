using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelCake.Intrusive;

namespace ExcelCake.Example
{
    [ExportModal("用户信息")]
    public class UserInfo: ExcelBase
    {
        [DisplayName("编号")]
        public int ID { set; get; }

        [DisplayName("姓名")]
        public string Name { set; get; }

        [DisplayName("性别")]
        public string Sex { set; get; }

        [DisplayName("年龄")]
        public int Age { set; get; }

        [DisplayName("电子邮件")]
        public string Email { set; get; }

        [DisplayName("联系方式")]
        public string TelPhone { set; get; }
    }
}
