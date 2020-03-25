using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace sforceAddin.sforce
{
    static class Util
    {
        public static string Error2String(this SFDC.Error error)
        {
            // return string.Format("Status: {1}, error message: {2}, fields: {0}, extended detail: {3}", error.fields, error.statusCode, error.message, error.extendedErrorDetails);
            return string.Format("{0}: {1}", error.statusCode, error.message);
        }

        public static string Error2String(this SFDC.Error[] errors)
        {
            StringBuilder sb = new StringBuilder();
            foreach (var err in errors)
            {
                sb.AppendLine(err.Error2String());
            }

            return sb.ToString();
        }

        public static int WordCount(this String str)
        {
            return str.Split(new char[] { ' ', '.', '?' },
                             StringSplitOptions.RemoveEmptyEntries).Length;
        }
    }
}
