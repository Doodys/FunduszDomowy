using System.Text;

namespace FunduszDomowy {
    class Utf8Converter {
        public string Utf8Convert(string mess) {
            byte[] bytes = Encoding.Default.GetBytes(mess);
            mess = Encoding.UTF8.GetString(bytes);
            return mess;
        }
    }
}
