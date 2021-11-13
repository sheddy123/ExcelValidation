using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelValidator.Helpers
{
    public static class Helpers
    {
        #region PR#7 Convert the first character of input string to capital
        public static string UpperCaseFirst(string inputString)
        {
            if (string.IsNullOrEmpty(inputString))
                return String.Empty;
            return char.ToUpper(inputString[0]) + inputString.Substring(1);
        }
        #endregion
    }
}
