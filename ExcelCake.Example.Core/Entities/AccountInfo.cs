using ExcelCake.Intrusive;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelCake.Example
{
    [ExportEntity("账号信息")]
    public class AccountInfo:ExcelBase
    {
        [Export("编号", 1)]
        public int ID { set; get; }

        [Export("昵称", 2)]
        public string Nickname { set; get; }

        [Export("密码", 3)]
        public string Password { set; get; }

        [Export("旧密码", 4)]
        public string OldPassword { set; get; }

        [Export("状态", 5)]
        public int AccountStatus { set; get; }
    }
}
