using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace FunduszDomowy {
    class IpkoConverter {

        public string CutMessage(string message) {
            Regex regexIncome = new Regex(@"(\d+),(\d+)");
            Regex regexMinus = new Regex(@"-(\d+),(\d+)");

            Match matchIncome = regexIncome.Match(message);
            Match matchMinus = regexMinus.Match(message);

            if (matchMinus.Success) {
                return matchMinus.Value; //return string with minus value
            } else if (matchIncome.Success) {
                return matchIncome.Value; //return string with income value
            } else return "";
        }

        public List<string> CutDate(string date, string amount)
        {
            List<string> lData = new List<string>();
            string[] sDate = date.Split(new char[] { ' ' });

            lData.Add(amount);
            foreach (string mess in sDate) { lData.Add(mess); }

            return lData;
        }
    }
}
