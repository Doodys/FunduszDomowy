using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace FunduszDomowy {
    class Config {
        private static Dictionary<string, string> ReadConfig() {

            Dictionary<string, string> configList = new Dictionary<string, string>();
            string path = "../../config.txt";

            var query = (
                from line in File.ReadAllLines(path)
                let values = line.Split('=')
                select new { Key = values[0], Value = values[1] }
                );

            foreach (var kvp in query) {
                configList[kvp.Key] = kvp.Value;
            }

            return configList;
        }

        public static string AddValues(string name) {
            string value = "";

            foreach (KeyValuePair<string, string> pair in ReadConfig()) {
                if (pair.Key == name) value = pair.Value;
            }
            return value;
        }

        internal string Mail = AddValues("Mail");
        internal string Password = AddValues("Password");
        internal string Bank = AddValues("Bank");
        internal string ExcelLang = AddValues("ExcelLang");
        internal string SaveDir = AddValues("SaveDir");
    }
}
