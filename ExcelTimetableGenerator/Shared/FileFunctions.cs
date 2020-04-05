using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelTimetableGenerator.Shared
{
    public class FileFunctions
    {
        public static string MakeValidFileName(string path)
        {
            string invalidChars = System.Text.RegularExpressions.Regex.Escape(new string(System.IO.Path.GetInvalidFileNameChars()));
            string invalidRegStr = string.Format(@"([{0}]*\.+$)|([{0}]+)", invalidChars);

            return System.Text.RegularExpressions.Regex.Replace(path, invalidRegStr, "_").Replace(" ", "_");
        }

        public static string FormatHyperlink(string url)
        {
            string invalidChars = System.Text.RegularExpressions.Regex.Escape(new string(System.IO.Path.GetInvalidPathChars()));
            string invalidRegStr = string.Format(@"([{0}]*\.+$)|([{0}]+)", invalidChars);

            return System.Text.RegularExpressions.Regex.Replace(url, invalidRegStr, "_").Replace(" ", "_");
        }

        public static string ShortenString(string str, int maxLength)
        {
            int strLength = str.Length;

            if (maxLength <= 3)
            {
                return str;
            }
            if (strLength > maxLength)
            {
                return str.Substring(0, maxLength - 3) + "...";
            }
            else
            {
                return str;
            }
        }
    }
}
