using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace ExcelCake
{
    [Serializable]
    public class ImportFormatException: ApplicationException
    {
        private string[] _Messages;

        public string[] Messages
        {
            get
            {
                return _Messages;
            }
        }

        public ImportFormatException()
        {
            _Messages = new string[0];
        }

        public ImportFormatException(params string[] messages)
        {
            _Messages = messages;
        }
    }
}