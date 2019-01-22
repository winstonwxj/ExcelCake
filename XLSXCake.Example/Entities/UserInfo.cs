using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XLSXCake.Intrusive;

namespace XLSXCake.Example
{
    [ExportModal(EnumExportModalType.INVERT, "用户信息")]
    public class UserInfo: XLSXBase
    {
        [DisplayName("编号")]
        [ExportSort(1)]
        public int ID { set; get; }

        [DisplayName("姓名")]
        [ExportSort(5)]
        public string Name { set; get; }

        [DisplayName("性别")]
        [ExportSort(4)]
        public string Sex { set; get; }

        [DisplayName("年龄")]
        [ExportSort(3)]
        public int Age { set; get; }

        [DisplayName("电子邮件")]
        [ExportSort(2)]
        public string Email { set; get; }

        [DisplayName("联系方式")]
        [NoExport]
        public string TelPhone { set; get; }
    }
}
