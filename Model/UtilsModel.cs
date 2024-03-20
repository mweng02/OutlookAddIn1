using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace OutlookAddIn1
{
    public class UtilsModel
    {

        public static string ExtractRegex(string Subject, string Pattern)
        {
            var match = Regex.Match(Subject, Pattern);
            if (match.Success)
            {
                return match.Value;
            }

            return string.Empty;
        }

    }
}
