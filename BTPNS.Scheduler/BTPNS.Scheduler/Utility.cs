using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace BTPNS.Scheduler
{
    public class Utility
    {
        public string GetStringValue(DataRow value, string key)
        {
            string result = string.Empty;
            try
            {
                return value[key].ToString();
            }
            catch
            {
                return result;
            }
        }

        public int GetIntValue(DataRow value, string key)
        {
            try
            {
                string val = value[key].ToString().Split('.')[0];
                return int.Parse(val);
            }
            catch
            {
                return 0;
            }
        }
        public decimal GetDecimalValue(DataRow value, string key)
        {
            try
            {
                string val = value[key].ToString();
                return Convert.ToDecimal(val);
            }
            catch
            {
                return 0;
            }
        }

        public DateTime GetDateValue(DataRow value, string key)
        {
            DateTime date = new DateTime();
            try
            {
                return DateTime.Parse(value[key].ToString());
            }
            catch
            {
                return date;
            }
        }

        public string StripHTML(string input)
        {
            return Regex.Replace(input, "<.*?>", String.Empty).Replace("&nbsp;", " ").Replace("&amp;", "&");
        }

        public string GetUntilOrEmpty(string text, string stopAt = "", string orStopAt = "")
        {
            if (!String.IsNullOrWhiteSpace(text))
            {
                int charLocation = text.IndexOf(stopAt, StringComparison.Ordinal);
                int charLocation2 = text.IndexOf(orStopAt, StringComparison.Ordinal);


                if (charLocation > 0)
                {
                    return text.Substring(0, charLocation);
                }
            }

            return String.Empty;
        }

        public static string ToHtmlTable(DataTable dt)
        {
            string strHtml = "<table><tr>" + Environment.NewLine;
            foreach (DataColumn col in dt.Columns)
            {
                strHtml += Environment.NewLine + "<th>" + col.ColumnName + "</th>";
            }
            strHtml += Environment.NewLine + "</tr>";
            foreach (DataColumn dc in dt.Columns)
            {
                strHtml += Environment.NewLine + "<tr>";
                foreach (DataRow row in dt.Rows)
                {
                    strHtml += Environment.NewLine + "<td>" + row[dc] + "</td>";
                }
                strHtml += Environment.NewLine + "</tr>";
            }
            strHtml += "</table>";
            return strHtml;
        }
    }
}
