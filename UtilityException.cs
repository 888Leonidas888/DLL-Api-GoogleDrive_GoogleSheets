using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GoogleServices
{
    public class UtilityException : Exception
    {
        public UtilityException(string message, int errorCode)
            : base(message)
        {
            HResult = errorCode;
        }
    }
}
